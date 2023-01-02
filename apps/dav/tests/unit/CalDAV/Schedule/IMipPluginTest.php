<?php
/**
 * @copyright Copyright (c) 2016, ownCloud, Inc.
 * @copyright Copyright (c) 2017, Georg Ehrke
 *
 * @author brad2014 <brad2014@users.noreply.github.com>
 * @author Brad Rubenstein <brad@wbr.tech>
 * @author Christoph Wurst <christoph@winzerhof-wurst.at>
 * @author Georg Ehrke <oc.list@georgehrke.com>
 * @author Joas Schilling <coding@schilljs.com>
 * @author Morris Jobke <hey@morrisjobke.de>
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
namespace OCA\DAV\Tests\unit\CalDAV\Schedule;

use OCA\DAV\CalDAV\Schedule\IMipPlugin;
use OCA\DAV\CalDAV\Schedule\IMipService;
use OCP\AppFramework\Utility\ITimeFactory;
use OCP\DB\QueryBuilder\IQueryBuilder;
use OCP\Defaults;
use OCP\IConfig;
use OCP\IDBConnection;
use OCP\IL10N;
use OCP\IURLGenerator;
use OCP\IUser;
use OCP\IUserManager;
use OCP\Mail\IAttachment;
use OCP\Mail\IEMailTemplate;
use OCP\Mail\IMailer;
use OCP\Mail\IMessage;
use PHPUnit\Framework\MockObject\MockObject;
use Psr\Log\LoggerInterface;
use Sabre\VObject\Component\VCalendar;
use Sabre\VObject\Component\VEvent;
use Sabre\VObject\ITip\Message;
use Sabre\VObject\Reader;
use Test\TestCase;
use function array_merge;

class IMipPluginTest extends TestCase {

	/** @var IMessage|MockObject */
	private $mailMessage;

	/** @var IMailer|MockObject */
	private $mailer;

	/** @var IEMailTemplate|MockObject */
	private $emailTemplate;

	/** @var IAttachment|MockObject */
	private $emailAttachment;

	/** @var ITimeFactory|MockObject */
	private $timeFactory;

	/** @var IConfig|MockObject */
	private $config;

	/** @var IUserManager|MockObject */
	private $userManager;

	/** @var IMipPlugin */
	private $plugin;

	/** @var IMipService|MockObject */
	private $service;

	/** @var Defaults|MockObject */
	private $defaults;

	/** @var LoggerInterface|MockObject */
	private $logger;

	protected function setUp(): void {
		$this->mailMessage = $this->createMock(IMessage::class);
		$this->mailMessage->method('setFrom')->willReturn($this->mailMessage);
		$this->mailMessage->method('setReplyTo')->willReturn($this->mailMessage);
		$this->mailMessage->method('setTo')->willReturn($this->mailMessage);

		$this->mailer = $this->createMock(IMailer::class);
		$this->mailer->method('createMessage')->willReturn($this->mailMessage);

		$this->emailTemplate = $this->createMock(IEMailTemplate::class);
		$this->mailer->method('createEMailTemplate')->willReturn($this->emailTemplate);

		$this->emailAttachment = $this->createMock(IAttachment::class);
		$this->mailer->method('createAttachment')->willReturn($this->emailAttachment);

		$this->logger = $this->createMock(LoggerInterface::class);

		$this->timeFactory = $this->createMock(ITimeFactory::class);
		$this->timeFactory->method('getTime')->willReturn(1496912528); // 2017-01-01

		$this->config = $this->createMock(IConfig::class);

		$this->userManager = $this->createMock(IUserManager::class);

		$this->defaults = $this->createMock(Defaults::class);
		$this->defaults->method('getName')
			->willReturn('Instance Name 123');

		$this->service = $this->createMock(IMipService::class);

		$this->plugin = new IMipPlugin(
			$this->config,
			$this->mailer,
			$this->logger,
			$this->timeFactory,
			$this->defaults,
			$this->userManager,
			'user123',
			$this->service);
	}

	public function testDeliveryNoSignificantChange(): void {
		$message = new Message();
		$message->method = 'REQUEST';
		$message->message = new VCalendar();
		$message->message->add('VEVENT', array_merge([
			'UID' => 'uid-1234',
			'SEQUENCE' => 0,
			'SUMMARY' => 'Fellowship meeting',
			'DTSTART' => new \DateTime('2016-01-01 00:00:00')
		], []));
		$message->message->VEVENT->add('ORGANIZER', 'mailto:gandalf@wiz.ard');
		$message->message->VEVENT->add('ATTENDEE', 'mailto:' . 'frodo@hobb.it', ['RSVP' => 'TRUE']);
		$message->sender = 'mailto:gandalf@wiz.ard';
		$message->senderName = 'Mr. Wizard';
		$message->recipient = 'mailto:' . 'frodo@hobb.it';
		$message->significantChange = false;
		$this->plugin->schedule($message);
		$this->assertEquals('1.0', $message->getScheduleStatus());
	}

	public function testParsingSingle(): void {
		$message = new Message();
		$message->method = 'REQUEST';
		$message->message = new VCalendar();
		$message->message->add('VEVENT', array_merge([
			'UID' => 'uid-1234',
			'SEQUENCE' => 1,
			'SUMMARY' => 'Fellowship meeting without (!) Boromir',
			'DTSTART' => new \DateTime('2016-01-01 00:00:00')
		], []));
		$message->message->VEVENT->add('ORGANIZER', 'mailto:gandalf@wiz.ard');
		$message->message->VEVENT->add('ATTENDEE', 'mailto:' . 'frodo@hobb.it', ['RSVP' => 'TRUE',  'CN' => 'Frodo']);
		$message->sender = 'mailto:gandalf@wiz.ard';
		$message->senderName = 'Mr. Wizard';
		$message->recipient = 'mailto:' . 'frodo@hobb.it';
		// save the old copy in the plugin
		$vCalendar = new VCalendar();
		$vEvent = new VEvent($vCalendar, 'one', [
			'UID' => 'uid-1234',
			'SEQUENCE' => 0,
			'SUMMARY' => 'Fellowship meeting',
			'DTSTART' => new \DateTime('2016-01-01 00:00:00')
		]);
		$vEvent->add('ORGANIZER', 'mailto:gandalf@wiz.ard');
		$vEvent->add('ATTENDEE', 'mailto:' . 'frodo@hobb.it', ['RSVP' => 'TRUE', 'CN' => 'Frodo']);
		$vEvent->add('ATTENDEE', 'mailto:' . 'boromir@tra.it.or', ['RSVP' => 'TRUE']);
		$vCalendar->add($vEvent);
		$data = ['invitee_name' => 'Mr. Wizard',
			'meeting_title' => 'Fellowship meeting without (!) Boromir',
			'attendee_name' => 'frodo@hobb.it'
		];
		$this->plugin->setVCalendar($vCalendar);
		$this->service->expects(self::once())
			->method('getLastOccurrence')
			->willReturn('1496912700');
		$this->mailer->expects(self::once())
			->method('validateMailAddress')
			->with('frodo@hobb.it')
			->willReturn(true);
		$this->service->expects(self::once())
			->method('processUnmodifieds')
			->willReturn(false);
		$this->service->expects(self::once())
			->method('buildBodyData')
			->with($message->message->VEVENT, $vEvent)
			->willReturn($data);
		$this->userManager->expects(self::once())
			->method('getDisplayName')
			->willReturn('Mr. Wizard');
		$this->service->expects(self::once())
			->method('getFrom');
		$this->service->expects(self::once())
			->method('addSubjectAndHeading')
			->with($this->emailTemplate, 'request', 'Mr. Wizard', 'Fellowship meeting without (!) Boromir', null);
		$this->service->expects(self::once())
			->method('addBulletList')
			->with($this->emailTemplate, $message->message->VEVENT, $data);
		$this->service->expects(self::once())
			->method('getAttendeeRsvpOrReqForParticipant')
			->willReturn(true);
		$this->config->expects(self::once())
			->method('getAppValue')
			->with('dav', 'invitation_link_recipients', 'yes')
			->willReturn('yes');
		$this->service->expects(self::once())
			->method('createInvitationToken')
			->with($message, $message->message->VEVENT, '1496912700')
			->willReturn('token');
		$this->service->expects(self::once())
			->method('addResponseButtons')
			->with($this->emailTemplate, 'token');
		$this->service->expects(self::once())
			->method('addMoreOptionsButton')
			->with($this->emailTemplate, 'token');
		$this->mailer->expects(self::once())
			->method('send')
			->willReturn([]);
		$this->plugin->schedule($message);
		$this->assertEquals('1.1', $message->getScheduleStatus());
	}

	public function testParsingRecurrence(): void {
		$message = new Message();
		$message->method = 'REQUEST';
		$message->message = new VCalendar();
		$message->message->add('VEVENT', [
			'UID' => 'uid-1234',
			'LAST-MODIFIED' => 123456,
			'SEQUENCE' => 2,
			'SUMMARY' => 'Fellowship meeting',
			'DTSTART' => new \DateTime('2016-01-01 00:00:00'),
			'RRULE' => 'FREQ=DAILY;INTERVAL=1;UNTIL=20160201T000000Z'
		]);
		$message->message->VEVENT->add('ORGANIZER', 'mailto:gandalf@wiz.ard');
		$message->message->VEVENT->add('ATTENDEE', 'mailto:' . 'frodo@hobb.it', ['RSVP' => 'TRUE',  'CN' => 'Frodo']);
		$message->sender = 'mailto:gandalf@wiz.ard';
		$message->senderName = 'Mr. Wizard';
		$message->recipient = 'mailto:' . 'frodo@hobb.it';
		// save the old copy in the plugin
		$vCalendar = new VCalendar();
		$vEvent = new VEvent($vCalendar, 'one', [
			'UID' => 'uid-1234',
			'LAST-MODIFIED' => 123456,
			'SEQUENCE' => 2,
			'SUMMARY' => 'Fellowship meeting',
			'DTSTART' => new \DateTime('2016-01-01 00:00:00'),
			'RRULE' => 'FREQ=DAILY;INTERVAL=1;UNTIL=20160201T000000Z'
		]);
		$vEvent->add('ORGANIZER', 'mailto:gandalf@wiz.ard');
		$vEvent->add('ATTENDEE', 'mailto:' . 'frodo@hobb.it', ['RSVP' => 'TRUE', 'CN' => 'Frodo']);
		$vCalendar->add($vEvent);
		$vEvent = new VEvent($vCalendar, 'two', [
			'UID' => 'uid-1234',
			'SEQUENCE' => 1,
			'SUMMARY' => 'Elevenses',
			'DTSTART' => new \DateTime('2016-01-01 00:00:00'),
			'RECURRENCE-ID' => new \DateTime('2016-01-01 00:00:00')
		]);
		$vEvent->add('ORGANIZER', 'mailto:gandalf@wiz.ard');
		$vEvent->add('ATTENDEE', 'mailto:' . 'frodo@hobb.it', ['RSVP' => 'TRUE', 'CN' => 'Frodo']);
		$vCalendar->add($vEvent);
		$data = ['invitee_name' => 'Mr. Wizard',
			'meeting_title' => 'Elevenses',
			'attendee_name' => 'frodo@hobb.it'
		];
		$this->plugin->setVCalendar($vCalendar);
		$this->service->expects(self::once())
			->method('getLastOccurrence')
			->willReturn('1496912700');
		$this->mailer->expects(self::once())
			->method('validateMailAddress')
			->with('frodo@hobb.it')
			->willReturn(true);
		$this->service->expects(self::exactly(2))
			->method('processUnmodifieds')
			->willReturnOnConsecutiveCalls(true, false);
		$this->service->expects(self::once())
			->method('buildBodyData')
			->with($message->message->VEVENT, $vEvent)
			->willReturn($data);
		$this->userManager->expects(self::once())
			->method('getDisplayName')
			->willReturn('Mr. Wizard');
		$this->service->expects(self::once())
			->method('getFrom');
		$this->service->expects(self::once())
			->method('addSubjectAndHeading')
			->with($this->emailTemplate, 'request', 'Mr. Wizard', 'Elevenses', null);
		$this->service->expects(self::once())
			->method('addBulletList')
			->with($this->emailTemplate, $message->message->VEVENT, $data);
		$this->service->expects(self::once())
			->method('getAttendeeRsvpOrReqForParticipant')
			->willReturn(true);
		$this->config->expects(self::once())
			->method('getAppValue')
			->with('dav', 'invitation_link_recipients', 'yes')
			->willReturn('yes');
		$this->service->expects(self::once())
			->method('createInvitationToken')
			->with($message, $message->message->VEVENT, '1496912700')
			->willReturn('token');
		$this->service->expects(self::once())
			->method('addResponseButtons')
			->with($this->emailTemplate, 'token');
		$this->service->expects(self::once())
			->method('addMoreOptionsButton')
			->with($this->emailTemplate, 'token');
		$this->mailer->expects(self::once())
			->method('send')
			->willReturn([]);
		$this->plugin->schedule($message);
		$this->assertEquals('1.1', $message->getScheduleStatus());
	}

	public function testEmailValidationFailed() {
		$message = new Message();
		$message->method = 'REQUEST';
		$message->message = new VCalendar();
		$message->message->add('VEVENT', array_merge([
			'UID' => 'uid-1234',
			'SEQUENCE' => 0,
			'SUMMARY' => 'Fellowship meeting',
			'DTSTART' => new \DateTime('2016-01-01 00:00:00')
		], []));
		$message->message->VEVENT->add('ORGANIZER', 'mailto:gandalf@wiz.ard');
		$message->message->VEVENT->add('ATTENDEE', 'mailto:' . 'frodo@hobb.it', ['RSVP' => 'TRUE']);
		$message->sender = 'mailto:gandalf@wiz.ard';
		$message->senderName = 'Mr. Wizard';
		$message->recipient = 'mailto:' . 'frodo@hobb.it';

		$this->service->expects(self::once())
			->method('getLastOccurrence')
			->willReturn('1496912700');
		$this->mailer->expects(self::once())
			->method('validateMailAddress')
			->with('frodo@hobb.it')
			->willReturn(false);

		$this->plugin->schedule($message);
		$this->assertEquals('5.0', $message->getScheduleStatus());
	}

	public function testFailedDelivery(): void {
		$message = new Message();
		$message->method = 'REQUEST';
		$message->message = new VCalendar();
		$message->message->add('VEVENT', array_merge([
			'UID' => 'uid-1234',
			'SEQUENCE' => 1,
			'SUMMARY' => 'Fellowship meeting without (!) Boromir',
			'DTSTART' => new \DateTime('2016-01-01 00:00:00')
		], []));
		$message->message->VEVENT->add('ORGANIZER', 'mailto:gandalf@wiz.ard');
		$message->message->VEVENT->add('ATTENDEE', 'mailto:' . 'frodo@hobb.it', ['RSVP' => 'TRUE',  'CN' => 'Frodo']);
		$message->sender = 'mailto:gandalf@wiz.ard';
		$message->senderName = 'Mr. Wizard';
		$message->recipient = 'mailto:' . 'frodo@hobb.it';
		// save the old copy in the plugin
		$vCalendar = new VCalendar();
		$vEvent = new VEvent($vCalendar, 'one', [
			'UID' => 'uid-1234',
			'SEQUENCE' => 0,
			'SUMMARY' => 'Fellowship meeting',
			'DTSTART' => new \DateTime('2016-01-01 00:00:00')
		]);
		$vEvent->add('ORGANIZER', 'mailto:gandalf@wiz.ard');
		$vEvent->add('ATTENDEE', 'mailto:' . 'frodo@hobb.it', ['RSVP' => 'TRUE', 'CN' => 'Frodo']);
		$vEvent->add('ATTENDEE', 'mailto:' . 'boromir@tra.it.or', ['RSVP' => 'TRUE']);
		$vCalendar->add($vEvent);
		$data = ['invitee_name' => 'Mr. Wizard',
			'meeting_title' => 'Fellowship meeting without (!) Boromir',
			'attendee_name' => 'frodo@hobb.it'
		];
		$this->plugin->setVCalendar($vCalendar);
		$this->service->expects(self::once())
			->method('getLastOccurrence')
			->willReturn('1496912700');
		$this->mailer->expects(self::once())
			->method('validateMailAddress')
			->with('frodo@hobb.it')
			->willReturn(true);
		$this->service->expects(self::once())
			->method('processUnmodifieds')
			->willReturn(false);
		$this->service->expects(self::once())
			->method('buildBodyData')
			->with($message->message->VEVENT, $vEvent)
			->willReturn($data);
		$this->userManager->expects(self::once())
			->method('getDisplayName')
			->willReturn('Mr. Wizard');
		$this->service->expects(self::once())
			->method('getFrom');
		$this->service->expects(self::once())
			->method('addSubjectAndHeading')
			->with($this->emailTemplate, 'request', 'Mr. Wizard', 'Fellowship meeting without (!) Boromir', null);
		$this->service->expects(self::once())
			->method('addBulletList')
			->with($this->emailTemplate, $message->message->VEVENT, $data);
		$this->service->expects(self::once())
			->method('getAttendeeRsvpOrReqForParticipant')
			->willReturn(true);
		$this->config->expects(self::once())
			->method('getAppValue')
			->with('dav', 'invitation_link_recipients', 'yes')
			->willReturn('yes');
		$this->service->expects(self::once())
			->method('createInvitationToken')
			->with($message, $message->message->VEVENT, '1496912700')
			->willReturn('token');
		$this->service->expects(self::once())
			->method('addResponseButtons')
			->with($this->emailTemplate, 'token');
		$this->service->expects(self::once())
			->method('addMoreOptionsButton')
			->with($this->emailTemplate, 'token');
		$this->mailer->expects(self::once())
			->method('send')
			->willReturn([]);
		$this->mailer
			->method('send')
			->willThrowException(new \Exception());
		$this->logger->expects(self::once())
			->method('error');
		$this->plugin->schedule($message);
		$this->assertEquals('5.0', $message->getScheduleStatus());
	}

	public function testNoOldEvent(): void {
		$message = new Message();
		$message->method = 'REQUEST';
		$message->message = new VCalendar();
		$message->message->add('VEVENT', array_merge([
			'UID' => 'uid-1234',
			'SEQUENCE' => 1,
			'SUMMARY' => 'Fellowship meeting',
			'DTSTART' => new \DateTime('2016-01-01 00:00:00')
		], []));
		$message->message->VEVENT->add('ORGANIZER', 'mailto:gandalf@wiz.ard');
		$message->message->VEVENT->add('ATTENDEE', 'mailto:' . 'frodo@hobb.it', ['RSVP' => 'TRUE',  'CN' => 'Frodo']);
		$message->sender = 'mailto:gandalf@wiz.ard';
		$message->senderName = 'Mr. Wizard';
		$message->recipient = 'mailto:' . 'frodo@hobb.it';
		$data = ['invitee_name' => 'Mr. Wizard',
			'meeting_title' => 'Fellowship meeting',
			'attendee_name' => 'frodo@hobb.it'
		];

		$this->service->expects(self::once())
			->method('getLastOccurrence')
			->willReturn('1496912700');
		$this->mailer->expects(self::once())
			->method('validateMailAddress')
			->with('frodo@hobb.it')
			->willReturn(true);
		$this->service->expects(self::never())
			->method('processUnmodifieds');
		$this->service->expects(self::once())
			->method('buildBodyData')
			->with($message->message->VEVENT, null)
			->willReturn($data);
		$this->userManager->expects(self::once())
			->method('getDisplayName')
			->willReturn('Mr. Wizard');
		$this->service->expects(self::once())
			->method('getFrom');
		$this->service->expects(self::once())
			->method('addSubjectAndHeading')
			->with($this->emailTemplate, 'request', 'Mr. Wizard', 'Fellowship meeting', null);
		$this->service->expects(self::once())
			->method('addBulletList')
			->with($this->emailTemplate, $message->message->VEVENT, $data);
		$this->service->expects(self::once())
			->method('getAttendeeRsvpOrReqForParticipant')
			->willReturn(true);
		$this->config->expects(self::once())
			->method('getAppValue')
			->with('dav', 'invitation_link_recipients', 'yes')
			->willReturn('yes');
		$this->service->expects(self::once())
			->method('createInvitationToken')
			->with($message, $message->message->VEVENT, '1496912700')
			->willReturn('token');
		$this->service->expects(self::once())
			->method('addResponseButtons')
			->with($this->emailTemplate, 'token');
		$this->service->expects(self::once())
			->method('addMoreOptionsButton')
			->with($this->emailTemplate, 'token');
		$this->mailer->expects(self::once())
			->method('send')
			->willReturn([]);
		$this->mailer
			->method('send')
			->willReturn([]);
		$this->plugin->schedule($message);
		$this->assertEquals('1.1', $message->getScheduleStatus());
	}

	public function testNoButtons(): void {
		$message = new Message();
		$message->method = 'REQUEST';
		$message->message = new VCalendar();
		$message->message->add('VEVENT', array_merge([
			'UID' => 'uid-1234',
			'SEQUENCE' => 1,
			'SUMMARY' => 'Fellowship meeting',
			'DTSTART' => new \DateTime('2016-01-01 00:00:00')
		], []));
		$message->message->VEVENT->add('ORGANIZER', 'mailto:gandalf@wiz.ard');
		$message->message->VEVENT->add('ATTENDEE', 'mailto:' . 'frodo@hobb.it', ['RSVP' => 'TRUE',  'CN' => 'Frodo']);
		$message->sender = 'mailto:gandalf@wiz.ard';
		$message->senderName = 'Mr. Wizard';
		$message->recipient = 'mailto:' . 'frodo@hobb.it';
		$data = ['invitee_name' => 'Mr. Wizard',
			'meeting_title' => 'Fellowship meeting',
			'attendee_name' => 'frodo@hobb.it'
		];

		$this->service->expects(self::once())
			->method('getLastOccurrence')
			->willReturn('1496912700');
		$this->mailer->expects(self::once())
			->method('validateMailAddress')
			->with('frodo@hobb.it')
			->willReturn(true);
		$this->service->expects(self::never())
			->method('processUnmodifieds');
		$this->service->expects(self::once())
			->method('buildBodyData')
			->with($message->message->VEVENT, null)
			->willReturn($data);
		$this->userManager->expects(self::once())
			->method('getDisplayName')
			->willReturn('Mr. Wizard');
		$this->service->expects(self::once())
			->method('getFrom');
		$this->service->expects(self::once())
			->method('addSubjectAndHeading')
			->with($this->emailTemplate, 'request', 'Mr. Wizard', 'Fellowship meeting', null);
		$this->service->expects(self::once())
			->method('addBulletList')
			->with($this->emailTemplate, $message->message->VEVENT, $data);
		$this->service->expects(self::once())
			->method('getAttendeeRsvpOrReqForParticipant')
			->willReturn(true);
		$this->config->expects(self::once())
			->method('getAppValue')
			->with('dav', 'invitation_link_recipients', 'yes')
			->willReturn('no');
		$this->service->expects(self::never())
			->method('createInvitationToken');
		$this->service->expects(self::never())
			->method('addResponseButtons');
		$this->service->expects(self::never())
			->method('addMoreOptionsButton');
		$this->mailer->expects(self::once())
			->method('send')
			->willReturn([]);
		$this->mailer
			->method('send')
			->willReturn([]);
		$this->plugin->schedule($message);
		$this->assertEquals('1.1', $message->getScheduleStatus());
	}
}
