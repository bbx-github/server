<?php

namespace OC\Share20;

use OCP\Server;
use OCP\Share\IShare;
use OCP\Share\IShareDisplayTemplateFactory;
use OCP\Share\IShareDisplayTemplateProvider;

class ShareDisplayTemplateFactory implements IShareDisplayTemplateFactory {
	/**
	 * @var IShareDisplayTemplateProvider[] $displayShareTemplateProviders
	 */
	private array $displayShareTemplateProviders = [];

	public function registerDisplayShareTemplate(string $shareDisplayTemplateClass): void {
		$this->displayShareTemplateProviders[] = $shareDisplayTemplateClass;
	}

	public function getTemplateProvider(IShare $share): IShareDisplayTemplateProvider {
		/**
		 * @var IShareDisplayTemplateProvider[]
		 */
		$providers = array_map(
			fn ($providerClass) => Server::get($providerClass),
			$this->displayShareTemplateProviders
		);
		usort($providers, fn (IShareDisplayTemplateProvider $a, IShareDisplayTemplateProvider $b) => $b->getPriority() - $a->getPriority());
		$filteredProviders = array_filter($providers, fn (IShareDisplayTemplateProvider $provider) => $provider->shouldRespond($share));
		return array_shift($filteredProviders);
	}
}
