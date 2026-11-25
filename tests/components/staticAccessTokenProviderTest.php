<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests\components;

use andrewsauder\microsoftServices\components\staticAccessTokenProvider;
use PHPUnit\Framework\Attributes\CoversClass;
use PHPUnit\Framework\Attributes\Test;
use PHPUnit\Framework\TestCase;

#[CoversClass( staticAccessTokenProvider::class )]
final class staticAccessTokenProviderTest extends TestCase {

	#[Test]
	public function itGivesTheTokenToMicrosoftGraph(): void {
		$provider = new staticAccessTokenProvider( 'token-1' );

		$this->assertSame( 'token-1', $provider->getAuthorizationTokenAsync( 'https://graph.microsoft.com/v1.0/users' )->wait() );
	}


	#[Test]
	public function itGivesNoTokenToAnotherHost(): void {
		$provider = new staticAccessTokenProvider( 'token-1' );

		$this->assertNull( $provider->getAuthorizationTokenAsync( 'https://contoso.sharepoint.com/upload/1' )->wait() );
	}

}
