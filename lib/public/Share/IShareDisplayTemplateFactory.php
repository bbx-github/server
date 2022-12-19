<?php

namespace OCP\Share;

interface IShareDisplayTemplateFactory {
	/**
	 * Register new display share template
	 * @param class-string<IShareDisplayTemplateProvider> $shareDisplayTemplateClass
	 */
	public function registerDisplayShareTemplate(string $shareDisplayTemplateClass): void;

	public function getTemplateProvider(IShare $share): IShareDisplayTemplateProvider;
}
