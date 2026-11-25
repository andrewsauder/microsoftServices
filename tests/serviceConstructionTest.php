<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests;

use andrewsauder\microsoftServices\calendar;
use andrewsauder\microsoftServices\config;
use andrewsauder\microsoftServices\exceptions\serviceException;
use andrewsauder\microsoftServices\files;
use andrewsauder\microsoftServices\mail;
use andrewsauder\microsoftServices\user;
use PHPUnit\Framework\Attributes\CoversClass;
use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\Attributes\Test;
use PHPUnit\Framework\TestCase;

/**
 * Every Graph service extends components\service, whose constructor resolves an access token via auth, which
 * first validates the config. So an invalid config is rejected before any network call is attempted - which is
 * what makes these constructors unit-testable offline.
 */
#[CoversClass( user::class )]
#[CoversClass( mail::class )]
#[CoversClass( calendar::class )]
#[CoversClass( files::class )]
final class serviceConstructionTest extends TestCase {

	#[Test]
	#[DataProvider( 'serviceClassProvider' )]
	public function constructingAServiceWithAnInvalidConfigThrowsBeforeAnyNetworkCall( string $serviceClass ): void {
		$this->expectException( serviceException::class );
		$this->expectExceptionCode( 400 );

		new $serviceClass( new config() );
	}


	/**
	 * The Graph client asks for a token only when it sends its first request, so a valid config builds every
	 * service offline: with an application token, with the on-behalf-of flow, and with a user token as it is.
	 */
	#[Test]
	#[DataProvider( 'authenticationModeProvider' )]
	public function constructingAServiceWithAValidConfigSendsNoRequest( bool $onBehalfOfFlow, ?string $userAccessToken ): void {
		$config                 = new config();
		$config->clientId       = 'client-id';
		$config->clientSecret   = 'client-secret';
		$config->tenant         = 'contoso.example';
		$config->driveId        = 'drive-1';
		$config->onBehalfOfFlow = $onBehalfOfFlow;

		foreach( self::serviceClassProvider() as [ $serviceClass ] ) {
			$this->assertInstanceOf( $serviceClass, new $serviceClass( $config, $userAccessToken ) );
		}
	}


	public static function authenticationModeProvider(): array {
		return [
			'application token' => [ true, null ],
			'on-behalf-of flow' => [ true, 'user-token' ],
			'user token'        => [ false, 'user-token' ],
		];
	}


	public static function serviceClassProvider(): array {
		return [
			'user'     => [ user::class ],
			'mail'     => [ mail::class ],
			'calendar' => [ calendar::class ],
			'files'    => [ files::class ],
		];
	}

}
