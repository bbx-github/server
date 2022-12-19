<?php

namespace OCP\Share;

use OCP\AppFramework\Http\TemplateResponse;

interface IShareDisplayTemplateProvider {
	public function getPriority(): int;
	public function shouldRespond(IShare $share): bool;
	public function renderPage(IShare $share, string $token, string $path): TemplateResponse;
}
