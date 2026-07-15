<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests;

use andrewsauder\microsoftServices\auth;
use andrewsauder\microsoftServices\config;
use andrewsauder\microsoftServices\exceptions\serviceException;
use PHPUnit\Framework\Attributes\CoversClass;
use PHPUnit\Framework\Attributes\Test;
use PHPUnit\Framework\TestCase;

#[CoversClass( auth::class )]
final class authTest extends TestCase {

	private function validConfig(): config {
		$config               = new config();
		$config->clientId     = 'client-id';
		$config->clientSecret = 'client-secret';
		$config->tenant       = 'tenant-id';

		return $config;
	}


	#[Test]
	public function constructorRejectsInvalidConfig(): void {
		//the auth constructor validates the config up front, before any Azure provider / network work
		$this->expectException( serviceException::class );
		$this->expectExceptionCode( 400 );
		new auth( new config() );
	}


	#[Test]
	public function constructorAcceptsValidConfigWithoutNetworkAccess(): void {
		//building the provider is purely local; no token exchange happens until getAccessToken()/verify()
		$auth = new auth( $this->validConfig() );
		$this->assertInstanceOf( auth::class, $auth );
	}

}
