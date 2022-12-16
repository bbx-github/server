<?php
/**
 * @copyright Copyright (c) 2016, ownCloud, Inc.
 * @copyright Copyright (c) 2017, Georg Ehrke
 * @copyright Copyright (C) 2007-2015 fruux GmbH (https://fruux.com/).
 * @copyright Copyright (C) 2007-2015 fruux GmbH (https://fruux.com/).
 *
 * @author brad2014 <brad2014@users.noreply.github.com>
 * @author Brad Rubenstein <brad@wbr.tech>
 * @author Christoph Wurst <christoph@winzerhof-wurst.at>
 * @author Georg Ehrke <oc.list@georgehrke.com>
 * @author Joas Schilling <coding@schilljs.com>
 * @author Leon Klingele <leon@struktur.de>
 * @author Nick Sweeting <git@sweeting.me>
 * @author rakekniven <mark.ziegler@rakekniven.de>
 * @author Roeland Jago Douma <roeland@famdouma.nl>
 * @author Thomas Citharel <nextcloud@tcit.fr>
 * @author Thomas Müller <thomas.mueller@tmit.eu>
 *
 * @license AGPL-3.0
 *
 * This code is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Affero General Public License, version 3,
 * as published by the Free Software Foundation.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Affero General Public License for more details.
 *
 * You should have received a copy of the GNU Affero General Public License, version 3,
 * along with this program. If not, see <http://www.gnu.org/licenses/>
 *
 */
namespace OCA\DAV\CalDAV\Schedule;

use OCA\DAV\CalDAV\CalendarObject;
use OCP\AppFramework\Utility\ITimeFactory;
use OCP\Defaults;
use OCP\IConfig;
use OCP\IDBConnection;
use OCP\IL10N;
use OCP\IURLGenerator;
use OCP\IUserManager;
use OCP\L10N\IFactory as L10NFactory;
use OCP\Mail\IEMailTemplate;
use OCP\Mail\IMailer;
use OCP\Security\ISecureRandom;
use OCP\Util;
use Psr\Log\LoggerInterface;
use Sabre\CalDAV\Schedule\IMipPlugin as SabreIMipPlugin;
use Sabre\DAV;
use Sabre\DAV\INode;
use Sabre\VObject\Component\VCalendar;
use Sabre\VObject\Component\VEvent;
use Sabre\VObject\Component\VTimeZone;
use Sabre\VObject\DateTimeParser;
use Sabre\VObject\ITip\Message;
use Sabre\VObject\Parameter;
use Sabre\VObject\Property;
use Sabre\VObject\Reader;
use Sabre\VObject\Recur\EventIterator;

/**
 * iMIP handler.
 *
 * This class is responsible for sending out iMIP messages. iMIP is the
 * email-based transport for iTIP. iTIP deals with scheduling operations for
 * iCalendar objects.
 *
 * If you want to customize the email that gets sent out, you can do so by
 * extending this class and overriding the sendMessage method.
 *
 * @copyright Copyright (C) 2007-2015 fruux GmbH (https://fruux.com/).
 * @author Evert Pot (http://evertpot.com/)
 * @license http://sabre.io/license/ Modified BSD License
 */
class IMipPlugin extends SabreIMipPlugin {

	/** @var string */
	private $userId;

	/** @var IConfig */
	private $config;

	/** @var IMailer */
	private $mailer;

	private LoggerInterface $logger;

	/** @var ITimeFactory */
	private $timeFactory;

	/** @var L10NFactory */
	private $l10nFactory;

	/** @var IURLGenerator */
	private $urlGenerator;

	/** @var ISecureRandom */
	private $random;

	/** @var IDBConnection */
	private $db;

	/** @var Defaults */
	private $defaults;

	/** @var IUserManager */
	private $userManager;

	public const MAX_DATE = '2038-01-01';
	public const METHOD_REQUEST = 'request';
	public const METHOD_REPLY = 'reply';
	public const METHOD_CANCEL = 'cancel';
	public const IMIP_INDENT = 15; // Enough for the length of all body bullet items, in all languages
	private ?VCalendar $vCalendar = null;

	public function __construct(IConfig $config, IMailer $mailer,
								LoggerInterface $logger,
								ITimeFactory $timeFactory, L10NFactory $l10nFactory,
								IURLGenerator $urlGenerator, Defaults $defaults,
								ISecureRandom $random, IDBConnection $db, IUserManager $userManager,
								$userId) {
		parent::__construct('');
		$this->userId = $userId;
		$this->config = $config;
		$this->mailer = $mailer;
		$this->logger = $logger;
		$this->timeFactory = $timeFactory;
		$this->l10nFactory = $l10nFactory;
		$this->urlGenerator = $urlGenerator;
		$this->random = $random;
		$this->db = $db;
		$this->defaults = $defaults;
		$this->userManager = $userManager;
	}

	public function initialize(DAV\Server $server) {
		parent::initialize($server);
		$server->on('beforeWriteContent', [$this, 'beforeWriteContent'], 10);
	}

	/**
	 * Check quota before writing content
	 *
	 * @param string $uri target file URI
	 * @param INode $node Sabre Node
	 * @param resource $data data
	 * @param bool $modified modified
	 */
	public function beforeWriteContent($uri, INode $node, $data, $modified) {
		if(!$node instanceof CalendarObject) {
			return;
		}
		/** @var VCalendar $vCalendar */
		$vCalendar = Reader::read($node->get());
		$this->setVCalendar($vCalendar);
	}

	private function processUnmodifieds(VEvent $event, array &$eventsToFilter): bool {
		/** @var VEvent $component */
		foreach ($eventsToFilter as $k => $component) {
			$componentRecurId = isset($component->{'RECURRENCE-ID'}) ? $component->{'RECURRENCE-ID'}->getValue() : null;
			$eventRecurId = isset($event->{'RECURRENCE-ID'}) ? $event->{'RECURRENCE-ID'}->getValue() : null;
			$componentRRule = isset($component->RRULE) ? $component->RRULE->getValue() : null;
			$eventRRule = isset($event->RRULE) ? $event->RRULE->getValue() : null;
			$componentSequence = isset($component->SEQUENCE) ? $component->SEQUENCE->getValue() : null;
			$eventSequence = isset($event->SEQUENCE) ? $event->SEQUENCE->getValue() : null;
			if(
				$component->{'LAST-MODIFIED'}->getValue() === $event->{'LAST-MODIFIED'}->getValue()
				&&$componentSequence === $eventSequence
				&& $componentRRule === $eventRRule
				&& $componentRecurId === $eventRecurId
			) {
				unset($eventsToFilter[$k]);
				return true;
			}
		}
		return false;
	}

	/**
	 * Event handler for the 'schedule' event.
	 *
	 * @param Message $iTipMessage
	 * @return void
	 */
	public function schedule(Message $iTipMessage) {
		// Not sending any emails if the system considers the update
		// insignificant.
		if (!$iTipMessage->significantChange) {
			if (!$iTipMessage->scheduleStatus) {
				$iTipMessage->scheduleStatus = '1.0;We got the message, but it\'s not significant enough to warrant an email';
			}
			return;
		}

		if (parse_url($iTipMessage->sender, PHP_URL_SCHEME) !== 'mailto' || parse_url($iTipMessage->recipient, PHP_URL_SCHEME) !== 'mailto') {
			return;
		}

		// don't send out mails for events that already took place
		$lastOccurrence = $this->getLastOccurrence($iTipMessage->message);
		$currentTime = $this->timeFactory->getTime();
		if ($lastOccurrence < $currentTime) {
			return;
		}

		// Strip off mailto:
		$recipient = substr($iTipMessage->recipient, 7);
		if ($recipient === false || !$this->mailer->validateMailAddress($recipient)) {
			// Nothing to send if the recipient doesn't have a valid email address
			$iTipMessage->scheduleStatus = '5.0; EMail delivery failed';
			return;
		}

		$newEvents = $iTipMessage->message;
		$newEventComponents = $newEvents->getComponents();

		foreach ($newEventComponents as $k => $event) {
			if($event instanceof VTimeZone) {
				unset($newEventComponents[$k]);
			}
		}

		$oldEvents = $this->getVCalendar();
		$oldEventComponents = $oldEvents === null ?: $oldEvents->getComponents();

		if(!is_array($oldEventComponents) || !empty($oldEventComponents)) {
			foreach ($oldEventComponents as $k => $event) {
				if($event instanceof VTimeZone) {
					unset($oldEventComponents[$k]);
					continue;
				}
				if($this->processUnmodifieds($event, $newEventComponents)) {
					unset($oldEventComponents[$k]);
				}
			}
		}

		// No changed events after all - this shouldn't happen if there is significant change yet here we are
		// The scheduling status is debatable
		// @todo handle this error case
		if(!is_array($newEventComponents) || empty($newEventComponents)) {
			$iTipMessage->scheduleStatus = '1.0;We got the message, but it\'s not significant enough to warrant an email';
			return;
		}

		// we (should) have one event component left
		// as the ITip\Broker creates one iTip message per change
		// and triggers the "schedule" event once per message
		// we also might not have an old event as this could be a new
		// invitation, or a new recurrence exception
		/** @var VEvent $vEvent */
		$vEvent = array_pop($newEventComponents);
		/** @var VEvent $oldVevent */
		$oldVevent = !empty($oldEventComponents) && is_array($oldEventComponents) ? array_pop($oldEventComponents) : null;
		$this->sendEmail($vEvent, $iTipMessage, $lastOccurrence, $oldVevent);
	}

	/**
	 * @param VEvent $vEvent
	 * @param IL10N $l10n
	 * @param VEvent|null $oldVEvent
	 * @return array
	 */
	private function buildBodyData(VEvent $vEvent, IL10N $l10n, ?VEvent $oldVEvent): array {
		$defaultVal = '';
		$strikethrough = "<span style='text-decoration: line-through'>%s</span><br />%s";

		$oldMeetingWhen = isset($oldVEvent) ? $this->generateWhenString($l10n, $oldVEvent) : null;
		$oldSummary = isset($oldVEvent->SUMMARY) && (string)$oldVEvent->SUMMARY !== '' ? (string)$oldVEvent->SUMMARY : $l10n->t('Untitled event');;
		$oldDescription = isset($oldVEvent->DESCRIPTION) && (string)$oldVEvent->DESCRIPTION !== '' ? (string)$oldVEvent->DESCRIPTION : $defaultVal;
		$oldUrl = isset($oldVEvent->URL) && (string)$oldVEvent->URL !== '' ? (string)$oldVEvent->URL : $defaultVal;
		$oldLocation = isset($oldVEvent->LOCATION) && (string)$oldVEvent->LOCATION !== '' ? (string)$oldVEvent->LOCATION : $defaultVal;

		$newMeetingWhen = $this->generateWhenString($l10n, $vEvent);
		$newSummary = isset($vEvent->SUMMARY) && (string)$vEvent->SUMMARY !== '' ? (string)$vEvent->SUMMARY : $l10n->t('Untitled event');;
		$newDescription = isset($vEvent->DESCRIPTION) && (string)$vEvent->DESCRIPTION !== '' ? (string)$vEvent->DESCRIPTION : $defaultVal;
		$newUrl = isset($vEvent->URL) && (string)$vEvent->URL !== '' ? sprintf('<a href="%1$s">%1$s</a>', $vEvent->URL) : $defaultVal;
		$newLocation = isset($vEvent->LOCATION) && (string)$vEvent->LOCATION !== ''  ? (string)$vEvent->LOCATION : $defaultVal;

		$data = [];
		$data['meeting_when'] = ($oldMeetingWhen !== $newMeetingWhen && $oldMeetingWhen !== null) ? sprintf($strikethrough, $oldMeetingWhen, $newMeetingWhen) : $newMeetingWhen;
		$data['meeting_when_plain'] = $newMeetingWhen;
		$data['meeting_title'] = ($oldSummary !== $newSummary) ? sprintf($strikethrough, $oldSummary, $newSummary) : $newSummary;
		$data['meeting_title_plain'] = $newSummary;
		$data['meeting_description'] = ($oldDescription !== $newDescription) ? sprintf($strikethrough, $oldDescription, $newDescription) : $newDescription;
		$data['meeting_description_plain'] = $newDescription;
		$data['meeting_url'] = ($oldUrl !== $newUrl) ? sprintf($strikethrough, $oldUrl, $newUrl) : $newUrl;
		$data['meeting_url_plain'] = isset($vEvent->URL) ? (string)$vEvent->URL : '';
		$data['meeting_location'] = ($oldLocation !== $newLocation) ? sprintf($strikethrough, $oldLocation, $newLocation) : $newLocation;
		$data['meeting_location_plain'] = $newLocation;
		return $data;
	}

	/**
	 * @param VEvent $vEvent
	 * @param IL10N $l10n
	 * @return array
	 */
	private function buildCancelledBodyData(VEvent $vEvent, IL10N $l10n): array {
		$defaultVal = '';
		$strikethrough = "<span style='text-decoration: line-through'>%$1s</span>";

		$newMeetingWhen = $this->generateWhenString($l10n, $vEvent);
		$newSummary = isset($vEvent->SUMMARY) && (string)$vEvent->SUMMARY !== '' ? (string)$vEvent->SUMMARY : $l10n->t('Untitled event');;
		$newDescription = isset($vEvent->DESCRIPTION) && (string)$vEvent->DESCRIPTION !== '' ? (string)$vEvent->DESCRIPTION : $defaultVal;
		$newUrl = isset($vEvent->URL) && (string)$vEvent->URL !== '' ? sprintf('<a href="%1$s">%1$s</a>', $vEvent->URL) : $defaultVal;
		$newLocation = isset($vEvent->LOCATION) && (string)$vEvent->LOCATION !== ''  ? (string)$vEvent->LOCATION : $defaultVal;

		$data = [];
		$data['meeting_when'] = $newMeetingWhen === '' ?: sprintf($strikethrough, $newMeetingWhen);
		$data['meeting_when_plain'] = $newMeetingWhen;
		$data['meeting_title'] = sprintf($strikethrough, $newSummary);
		$data['meeting_title_plain'] = $newSummary !== '' ?: $l10n->t('Untitled event');
		$data['meeting_description'] = $newDescription === '' ?: sprintf($strikethrough, $newDescription);
		$data['meeting_description_plain'] = $newDescription;
		$data['meeting_url'] =  $newUrl === '' ?:  sprintf($strikethrough, $newUrl);
		$data['meeting_url_plain'] = isset($vEvent->URL) ? (string)$vEvent->URL : '';
		$data['meeting_location'] = $newLocation === '' ?:  sprintf($strikethrough, $newLocation);
		$data['meeting_location_plain'] = $newLocation;
		return $data;
	}


	/**
	 * @param VEvent $vEvent
	 * @param Message $iTipMessage
	 * @param int $lastOccurrence
	 * @return void
	 */
	private function sendEmail(VEvent $vEvent, Message $iTipMessage, int $lastOccurrence, ?VEvent $oldVEvent = null): void {
		$attendee = $this->getCurrentAttendee($iTipMessage);
		$defaultLang = $this->l10nFactory->findGenericLanguage();
		$lang = $this->getAttendeeLangOrDefault($defaultLang, $attendee);
		$l10n = $this->l10nFactory->get('dav', $lang);

		switch (strtolower($iTipMessage->method)) {
			case self::METHOD_REPLY:
				$method = self::METHOD_REPLY;
				$data = $this->buildBodyData($vEvent, $l10n, $oldVEvent);
				break;
			case self::METHOD_CANCEL:
				$method = self::METHOD_CANCEL;
				$data = $this->buildCancelledBodyData($vEvent, $l10n);
				break;
			default:
				$method = self::METHOD_REQUEST;
				$data = $this->buildBodyData($vEvent, $l10n, $oldVEvent);
				break;
		}

		$recipient = substr($iTipMessage->recipient, 7);
		$recipientName = $iTipMessage->recipientName ?: null;

		$sender = substr($iTipMessage->sender, 7);
		$senderName = $iTipMessage->senderName ?: null;
		if ($senderName === null || empty(trim($senderName))) {
			$senderName = $this->userManager->getDisplayName($this->userId);
		}

		$data['attendee_name'] = ($recipientName ?: $recipient);
		$data['invitee_name'] = ($senderName ?: $sender);

		$fromEMail = Util::getDefaultEmailAddress('invitations-noreply');
		$fromName = $l10n->t('%1$s via %2$s', [$senderName ?? $this->userId, $this->defaults->getName()]);

		$message = $this->mailer->createMessage()
			->setFrom([$fromEMail => $fromName])
			->setTo([$recipient => $recipientName]);

		if ($sender !== false) {
			$message->setReplyTo([$sender => $senderName]);
		}

		$template = $this->mailer->createEMailTemplate('dav.calendarInvite.' . $method, $data);
		$template->addHeader();

		$this->addSubjectAndHeading($template, $l10n, $method, $data['invitee_name'], $data['meeting_title_plain']);
		$this->addBulletList($template, $l10n, $vEvent, $data);

		// Only add response buttons to invitation requests: Fix Issue #11230
		if (($method == self::METHOD_REQUEST) && $this->getAttendeeRsvpOrReqForParticipant($attendee)) {

			/*
			** Only offer invitation accept/reject buttons, which link back to the
			** nextcloud server, to recipients who can access the nextcloud server via
			** their internet/intranet.  Issue #12156
			**
			** The app setting is stored in the appconfig database table.
			**
			** For nextcloud servers accessible to the public internet, the default
			** "invitation_link_recipients" value "yes" (all recipients) is appropriate.
			**
			** When the nextcloud server is restricted behind a firewall, accessible
			** only via an internal network or via vpn, you can set "dav.invitation_link_recipients"
			** to the email address or email domain, or comma separated list of addresses or domains,
			** of recipients who can access the server.
			**
			** To always deliver URLs, set invitation_link_recipients to "yes".
			** To suppress URLs entirely, set invitation_link_recipients to boolean "no".
			*/

			$recipientDomain = substr(strrchr($recipient, '@'), 1);
			$invitationLinkRecipients = explode(',', preg_replace('/\s+/', '', strtolower($this->config->getAppValue('dav', 'invitation_link_recipients', 'yes'))));

			if (strcmp('yes', $invitationLinkRecipients[0]) === 0
				|| in_array(strtolower($recipient), $invitationLinkRecipients)
				|| in_array(strtolower($recipientDomain), $invitationLinkRecipients)) {
				$this->addResponseButtons($template, $l10n, $iTipMessage, $vEvent, $lastOccurrence);
			}
		}

		$template->addFooter();

		$message->useTemplate($template);

		// Let's clone all components of the iTip Calendar so we get a correct iMip Email
		// We only want the single component that was modifed as a VEVENT
		$vCalendar = new VCalendar();
		$vCalendar->add('METHOD', $iTipMessage->method);
		foreach ($iTipMessage->message->getComponents() as $component) {
			if($component instanceof VEvent) {
				continue;
			}
			$vCalendar->add(clone $component);
		}
		$vCalendar->add($vEvent);

		$attachment = $this->mailer->createAttachment(
			$vCalendar->serialize(),
			'event.ics',
			'text/calendar; method=' . $iTipMessage->method
		);
		$message->attach($attachment);

		try {
			$failed = $this->mailer->send($message);
			$iTipMessage->scheduleStatus = '1.1; Scheduling message is sent via iMip';
			if ($failed) {
				$this->logger->error('Unable to deliver message to {failed}', ['app' => 'dav', 'failed' => implode(', ', $failed)]);
				$iTipMessage->scheduleStatus = '5.0; EMail delivery failed';
			}
		} catch (\Exception $ex) {
			$this->logger->error($ex->getMessage(), ['app' => 'dav', 'exception' => $ex]);
			$iTipMessage->scheduleStatus = '5.0; EMail delivery failed';
		}
	}

	/**
	 * check if event took place in the past already
	 * @param VCalendar $vObject
	 * @return int
	 */
	private function getLastOccurrence(VCalendar $vObject) {
		/** @var VEvent $component */
		$component = $vObject->VEVENT;

		$firstOccurrence = $component->DTSTART->getDateTime()->getTimeStamp();
		// Finding the last occurrence is a bit harder
		if (!isset($component->RRULE)) {
			if (isset($component->DTEND)) {
				$lastOccurrence = $component->DTEND->getDateTime()->getTimeStamp();
			} elseif (isset($component->DURATION)) {
				/** @var \DateTime $endDate */
				$endDate = clone $component->DTSTART->getDateTime();
				// $component->DTEND->getDateTime() returns DateTimeImmutable
				$endDate = $endDate->add(DateTimeParser::parse($component->DURATION->getValue()));
				$lastOccurrence = $endDate->getTimestamp();
			} elseif (!$component->DTSTART->hasTime()) {
				/** @var \DateTime $endDate */
				$endDate = clone $component->DTSTART->getDateTime();
				// $component->DTSTART->getDateTime() returns DateTimeImmutable
				$endDate = $endDate->modify('+1 day');
				$lastOccurrence = $endDate->getTimestamp();
			} else {
				$lastOccurrence = $firstOccurrence;
			}
		} else {
			$it = new EventIterator($vObject, (string)$component->UID);
			$maxDate = new \DateTime(self::MAX_DATE);
			if ($it->isInfinite()) {
				$lastOccurrence = $maxDate->getTimestamp();
			} else {
				$end = $it->getDtEnd();
				while ($it->valid() && $end < $maxDate) {
					$end = $it->getDtEnd();
					$it->next();
				}
				$lastOccurrence = $end->getTimestamp();
			}
		}

		return $lastOccurrence;
	}

	/**
	 * @param Message $iTipMessage
	 * @return null|Property
	 */
	private function getCurrentAttendee(Message $iTipMessage) {
		/** @var VEvent $vevent */
		$vevent = $iTipMessage->message->VEVENT;
		$attendees = $vevent->select('ATTENDEE');
		foreach ($attendees as $attendee) {
			/** @var Property $attendee */
			if (strcasecmp($attendee->getValue(), $iTipMessage->recipient) === 0) {
				return $attendee;
			}
		}
		return null;
	}

	/**
	 * @param string $default
	 * @param Property|null $attendee
	 * @return string
	 */
	private function getAttendeeLangOrDefault($default, Property $attendee = null) {
		if ($attendee !== null) {
			$lang = $attendee->offsetGet('LANGUAGE');
			if ($lang instanceof Parameter) {
				return $lang->getValue();
			}
		}
		return $default;
	}

	/**
	 * @param Property|null $attendee
	 * @return bool
	 */
	private function getAttendeeRsvpOrReqForParticipant(Property $attendee = null) {
		if ($attendee !== null) {
			$rsvp = $attendee->offsetGet('RSVP');
			if (($rsvp instanceof Parameter) && (strcasecmp($rsvp->getValue(), 'TRUE') === 0)) {
				return true;
			}
			$role = $attendee->offsetGet('ROLE');
			// @see https://datatracker.ietf.org/doc/html/rfc5545#section-3.2.16
			// Attendees without a role are assumed required and should receive an invitation link even if they have no RSVP set
			if ($role === null
				|| (($role instanceof Parameter) && (strcasecmp($role->getValue(), 'REQ-PARTICIPANT') === 0))
				|| (($role instanceof Parameter) && (strcasecmp($role->getValue(), 'OPT-PARTICIPANT') === 0))
			) {
				return true;
			}
		}
		// RFC 5545 3.2.17: default RSVP is false
		return false;
	}

	/**
	 * @param IL10N $l10n
	 * @param VEvent $vevent
	 */
	private function generateWhenString(IL10N $l10n, VEvent $vevent) {
		$dtstart = $vevent->DTSTART;
		if (isset($vevent->DTEND)) {
			$dtend = $vevent->DTEND;
		} elseif (isset($vevent->DURATION)) {
			$isFloating = $vevent->DTSTART->isFloating();
			$dtend = clone $vevent->DTSTART;
			$endDateTime = $dtend->getDateTime();
			$endDateTime = $endDateTime->add(DateTimeParser::parse($vevent->DURATION->getValue()));
			$dtend->setDateTime($endDateTime, $isFloating);
		} elseif (!$vevent->DTSTART->hasTime()) {
			$isFloating = $vevent->DTSTART->isFloating();
			$dtend = clone $vevent->DTSTART;
			$endDateTime = $dtend->getDateTime();
			$endDateTime = $endDateTime->modify('+1 day');
			$dtend->setDateTime($endDateTime, $isFloating);
		} else {
			$dtend = clone $vevent->DTSTART;
		}

		$isAllDay = $dtstart instanceof Property\ICalendar\Date;

		/** @var Property\ICalendar\Date | Property\ICalendar\DateTime $dtstart */
		/** @var Property\ICalendar\Date | Property\ICalendar\DateTime $dtend */
		/** @var \DateTimeImmutable $dtstartDt */
		$dtstartDt = $dtstart->getDateTime();
		/** @var \DateTimeImmutable $dtendDt */
		$dtendDt = $dtend->getDateTime();

		$diff = $dtstartDt->diff($dtendDt);

		$dtstartDt = new \DateTime($dtstartDt->format(\DateTimeInterface::ATOM));
		$dtendDt = new \DateTime($dtendDt->format(\DateTimeInterface::ATOM));

		if ($isAllDay) {
			// One day event
			if ($diff->days === 1) {
				return $l10n->l('date', $dtstartDt, ['width' => 'medium']);
			}

			// DTEND is exclusive, so if the ics data says 2020-01-01 to 2020-01-05,
			// the email should show 2020-01-01 to 2020-01-04.
			$dtendDt->modify('-1 day');

			//event that spans over multiple days
			$localeStart = $l10n->l('date', $dtstartDt, ['width' => 'medium']);
			$localeEnd = $l10n->l('date', $dtendDt, ['width' => 'medium']);

			return $localeStart . ' - ' . $localeEnd;
		}

		/** @var Property\ICalendar\DateTime $dtstart */
		/** @var Property\ICalendar\DateTime $dtend */
		$isFloating = $dtstart->isFloating();
		$startTimezone = $endTimezone = null;
		if (!$isFloating) {
			$prop = $dtstart->offsetGet('TZID');
			if ($prop instanceof Parameter) {
				$startTimezone = $prop->getValue();
			}

			$prop = $dtend->offsetGet('TZID');
			if ($prop instanceof Parameter) {
				$endTimezone = $prop->getValue();
			}
		}

		$localeStart = $l10n->l('weekdayName', $dtstartDt, ['width' => 'abbreviated']) . ', ' .
			$l10n->l('datetime', $dtstartDt, ['width' => 'medium|short']);

		// always show full date with timezone if timezones are different
		if ($startTimezone !== $endTimezone) {
			$localeEnd = $l10n->l('datetime', $dtendDt, ['width' => 'medium|short']);

			return $localeStart . ' (' . $startTimezone . ') - ' .
				$localeEnd . ' (' . $endTimezone . ')';
		}

		// show only end time if date is the same
		if ($this->isDayEqual($dtstartDt, $dtendDt)) {
			$localeEnd = $l10n->l('time', $dtendDt, ['width' => 'short']);
		} else {
			$localeEnd = $l10n->l('weekdayName', $dtendDt, ['width' => 'abbreviated']) . ', ' .
				$l10n->l('datetime', $dtendDt, ['width' => 'medium|short']);
		}

		return  $localeStart . ' - ' . $localeEnd . ' (' . $startTimezone . ')';
	}

	/**
	 * @param \DateTime $dtStart
	 * @param \DateTime $dtEnd
	 * @return bool
	 */
	private function isDayEqual(\DateTime $dtStart, \DateTime $dtEnd) {
		return $dtStart->format('Y-m-d') === $dtEnd->format('Y-m-d');
	}

	/**
	 * @param IEMailTemplate $template
	 * @param IL10N $l10n
	 * @param string $method
	 * @param string $summary
	 */
	private function addSubjectAndHeading(IEMailTemplate $template, IL10N $l10n,
										  $method, Property $attendee, $sender, $summary) {
		if ($method === self::METHOD_CANCEL) {
			// TRANSLATORS Subject for email, when an invitation is cancelled. Ex: "Cancelled: {{Event Name}}"
			$template->setSubject($l10n->t('Cancelled: %1$s', [$summary]));
			$template->addHeading($l10n->t('"%1$s" has been canceled', [$summary]));
		} elseif ($method === self::METHOD_REPLY) {
			// TRANSLATORS Subject for email, when an invitation is replied to. Ex: "Re: {{Event Name}}"
			// Technically, the sender should be the attendee here (famous last words)
			switch (strtolower($attendee->offsetGet('PARTSTAT'))) {
				case 'accepted':
					$partstat = $l10n->t('%1$s has accepted your invitation', [$sender]);
					break;
				case 'tentative':
					$partstat = $l10n->t('%1$s has tentatively accepted your invitation', [$sender]);
					break;
				case 'declined':
					$partstat = $l10n->t('%1$s has declined your invitation', [$sender]);
					break;
				default:
					$partstat = $l10n->t('%1$s has responded your invitation', [$sender]);
					break;
			}
			$template->setSubject($l10n->t('Re: %1$s', [$summary]));
			$template->addHeading($partstat);
		} else {
			// TRANSLATORS Subject for email, when an invitation is sent. Ex: "Invitation: {{Event Name}}"
			$template->setSubject($l10n->t('Invitation: %1$s', [$summary]));
			$template->addHeading($l10n->t('%1$s would like to invite you to "%2$s"', [$sender, $summary]));
		}
	}

	/**
	 * @param IEMailTemplate $template
	 * @param IL10N $l10n
	 * @param VEVENT $vevent
	 */
	private function addBulletList(IEMailTemplate $template, IL10N $l10n, VEvent $vevent, $data) {
		$template->addBodyListItem(
			$data['meeting_title'], $l10n->t('Title:'),
			$this->getAbsoluteImagePath('caldav/title.png'), $data['meeting_title_plain'], '', self::IMIP_INDENT);
		if ($data['meeting_when'] !== '') {
			$template->addBodyListItem($data['meeting_when'], $l10n->t('Time:'),
				$this->getAbsoluteImagePath('caldav/time.png'),$data['meeting_when_plain'],'',self::IMIP_INDENT);
		}
		if ($data['meeting_location'] !== '') {
			$template->addBodyListItem($data['meeting_location'], $l10n->t('Location:'),
				$this->getAbsoluteImagePath('caldav/location.png'),$data['meeting_location_plain'],'',self::IMIP_INDENT);
		}
		if ($data['meeting_url'] !== '') {
			$template->addBodyListItem($data['meeting_url'], $l10n->t('Link:'),
				$this->getAbsoluteImagePath('caldav/link.png'), $data['meeting_url_plain'], '',self::IMIP_INDENT);
		}

		$this->addAttendees($template, $l10n, $vevent);

		/* Put description last, like an email body, since it can be arbitrarily long */
		if ($data['meeting_description']) {
			$template->addBodyListItem($data['meeting_description'], $l10n->t('Description:'),
				$this->getAbsoluteImagePath('caldav/description.png'),$data['meeting_description_plain'],'',self::IMIP_INDENT);
		}
	}

	/**
	 * addAttendees: add organizer and attendee names/emails to iMip mail.
	 *
	 * Enable with DAV setting: invitation_list_attendees (default: no)
	 *
	 * The default is 'no', which matches old behavior, and is privacy preserving.
	 *
	 * To enable including attendees in invitation emails:
	 *   % php occ config:app:set dav invitation_list_attendees --value yes
	 *
	 * @param IEMailTemplate $template
	 * @param IL10N $l10n
	 * @param Message $iTipMessage
	 * @param int $lastOccurrence
	 * @author brad2014 on github.com
	 */

	private function addAttendees(IEMailTemplate $template, IL10N $l10n, VEvent $vevent) {
		if ($this->config->getAppValue('dav', 'invitation_list_attendees', 'no') === 'no') {
			return;
		}

		if (isset($vevent->ORGANIZER)) {
			/** @var Property\ICalendar\CalAddress $organizer */
			$organizer = $vevent->ORGANIZER;
			$organizerURI = $organizer->getNormalizedValue();
			[$scheme,$organizerEmail] = explode(':',$organizerURI,2); # strip off scheme mailto:
			/** @var string|null $organizerName */
			$organizerName = isset($organizer['CN']) ? $organizer['CN'] : null;
			$organizerHTML = sprintf('<a href="%s">%s</a>',
				htmlspecialchars($organizerURI),
				htmlspecialchars($organizerName ?: $organizerEmail));
			$organizerText = sprintf('%s <%s>', $organizerName, $organizerEmail);
			if (isset($organizer['PARTSTAT'])) {
				/** @var Parameter $partstat */
				$partstat = $organizer['PARTSTAT'];
				if (strcasecmp($partstat->getValue(), 'ACCEPTED') === 0) {
					$organizerHTML .= ' ✔︎';
					$organizerText .= ' ✔︎';
				}
			}
			$template->addBodyListItem($organizerHTML, $l10n->t('Organizer:'),
				$this->getAbsoluteImagePath('caldav/organizer.png'),
				$organizerText,'',self::IMIP_INDENT);
		}

		$attendees = $vevent->select('ATTENDEE');
		if (count($attendees) === 0) {
			return;
		}

		$attendeesHTML = [];
		$attendeesText = [];
		foreach ($attendees as $attendee) {
			$attendeeURI = $attendee->getNormalizedValue();
			[$scheme,$attendeeEmail] = explode(':',$attendeeURI,2); # strip off scheme mailto:
			$attendeeName = isset($attendee['CN']) ? $attendee['CN'] : null;
			$attendeeHTML = sprintf('<a href="%s">%s</a>',
				htmlspecialchars($attendeeURI),
				htmlspecialchars($attendeeName ?: $attendeeEmail));
			$attendeeText = sprintf('%s <%s>', $attendeeName, $attendeeEmail);
			if (isset($attendee['PARTSTAT'])
				&& strcasecmp($attendee['PARTSTAT'], 'ACCEPTED') === 0) {
				$attendeeHTML .= ' ✔︎';
				$attendeeText .= ' ✔︎';
			}
			$attendeesHTML[] = $attendeeHTML;
			$attendeesText[] = $attendeeText;
		}

		$template->addBodyListItem(implode('<br/>',$attendeesHTML), $l10n->t('Attendees:'),
			$this->getAbsoluteImagePath('caldav/attendees.png'),
			implode("\n",$attendeesText),'',self::IMIP_INDENT);
	}

	/**
	 * @param IEMailTemplate $template
	 * @param IL10N $l10n
	 * @param Message $iTipMessage
	 * @param int $lastOccurrence
	 */
	private function addResponseButtons(IEMailTemplate $template, IL10N $l10n, Message $iTipMessage, VEvent $vevent, $lastOccurrence) {
		$token = $this->createInvitationToken($iTipMessage, $vevent, $lastOccurrence);

		$template->addBodyButtonGroup(
			$l10n->t('Accept'),
			$this->urlGenerator->linkToRouteAbsolute('dav.invitation_response.accept', [
				'token' => $token,
			]),
			$l10n->t('Decline'),
			$this->urlGenerator->linkToRouteAbsolute('dav.invitation_response.decline', [
				'token' => $token,
			])
		);

		$moreOptionsURL = $this->urlGenerator->linkToRouteAbsolute('dav.invitation_response.options', [
			'token' => $token,
		]);
		$html = vsprintf('<small><a href="%s">%s</a></small>', [
			$moreOptionsURL, $l10n->t('More options …')
		]);
		$text = $l10n->t('More options at %s', [$moreOptionsURL]);

		$template->addBodyText($html, $text);
	}

	/**
	 * @param string $path
	 * @return string
	 */
	private function getAbsoluteImagePath($path) {
		return $this->urlGenerator->getAbsoluteURL(
			$this->urlGenerator->imagePath('core', $path)
		);
	}

	/**
	 * @param Message $iTipMessage
	 * @param int $lastOccurrence
	 * @return string
	 */
	private function createInvitationToken(Message $iTipMessage, VEvent $vevent, $lastOccurrence):string {
		$token = $this->random->generate(60, ISecureRandom::CHAR_ALPHANUMERIC);

		$attendee = $iTipMessage->recipient;
		$organizer = $iTipMessage->sender;
		$sequence = $iTipMessage->sequence;
		$recurrenceId = isset($vevent->{'RECURRENCE-ID'}) ?
			$vevent->{'RECURRENCE-ID'}->serialize() : null;
		$uid = $vevent->{'UID'};

		$query = $this->db->getQueryBuilder();
		$query->insert('calendar_invitations')
			->values([
				'token' => $query->createNamedParameter($token),
				'attendee' => $query->createNamedParameter($attendee),
				'organizer' => $query->createNamedParameter($organizer),
				'sequence' => $query->createNamedParameter($sequence),
				'recurrenceid' => $query->createNamedParameter($recurrenceId),
				'expiration' => $query->createNamedParameter($lastOccurrence),
				'uid' => $query->createNamedParameter($uid)
			])
			->execute();

		return $token;
	}

	/**
	 * @return ?VCalendar
	 */
	public function getVCalendar(): ?VCalendar {
		return $this->vCalendar;
	}

	/**
	 * @param ?VCalendar $vCalendar
	 */
	public function setVCalendar(?VCalendar $vCalendar): void {
		$this->vCalendar = $vCalendar;
	}

}
