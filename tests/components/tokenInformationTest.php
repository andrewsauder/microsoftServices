<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests\components;

use andrewsauder\microsoftServices\components\tokenInformation;
use PHPUnit\Framework\Attributes\CoversClass;
use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\Attributes\Test;
use PHPUnit\Framework\TestCase;

#[CoversClass( tokenInformation::class )]
final class tokenInformationTest extends TestCase {

	#[Test]
	#[DataProvider( 'stringPropertyProvider' )]
	public function everyClaimPropertyDefaultsToEmptyString( string $property ): void {
		$tokenInformation = new tokenInformation();

		$this->assertSame( '', $tokenInformation->$property );
	}


	public static function stringPropertyProvider(): array {
		return [
			[ 'aud' ],
			[ 'iss' ],
			[ 'oid' ],
			[ 'sub' ],
			[ 'appId' ],
			[ 'name' ],
			[ 'familyName' ],
			[ 'givenName' ],
			[ 'ip' ],
			[ 'scope' ],
			[ 'email' ],
			[ 'preferred_username' ],
			[ 'upn' ],
			[ 'uniqueName' ],
		];
	}


	#[Test]
	public function propertiesAreAssignable(): void {
		$tokenInformation        = new tokenInformation();
		$tokenInformation->email = 'user@example.com';
		$tokenInformation->oid   = 'object-id';

		$this->assertSame( 'user@example.com', $tokenInformation->email );
		$this->assertSame( 'object-id', $tokenInformation->oid );
	}

}
