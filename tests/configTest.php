<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests;

use andrewsauder\microsoftServices\config;
use andrewsauder\microsoftServices\exceptions\serviceException;
use PHPUnit\Framework\Attributes\CoversClass;
use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\Attributes\Test;
use PHPUnit\Framework\TestCase;

#[CoversClass( config::class )]
final class configTest extends TestCase {

	/**
	 * Build a config that passes the base validate() checks.
	 */
	private function validConfig(): config {
		$config               = new config();
		$config->clientId     = 'client-id';
		$config->clientSecret = 'client-secret';
		$config->tenant       = 'tenant-id';

		return $config;
	}


	#[Test]
	public function defaultScopeIsPopulated(): void {
		$config = new config();
		$this->assertSame( 'openid profile email offline_access', $config->scope );
	}


	#[Test]
	public function onBehalfOfFlowDefaultsToTrue(): void {
		$config = new config();
		$this->assertTrue( $config->onBehalfOfFlow );
	}


	#[Test]
	public function freshConfigReportsMissingRequiredFields(): void {
		$config = new config();
		$errors = $config->validate( false );

		$this->assertSame( [
			'Client ID is required',
			'Client secret is required',
			'Tenant is required',
		], $errors );
	}


	#[Test]
	public function validConfigReturnsNoErrors(): void {
		$errors = $this->validConfig()->validate( false );
		$this->assertSame( [], $errors );
	}


	#[Test]
	public function validConfigDoesNotThrow(): void {
		$this->assertSame( [], $this->validConfig()->validate() );
	}


	#[Test]
	public function emptyScopeIsReported(): void {
		$config        = $this->validConfig();
		$config->scope = '';
		$this->assertSame( [ 'Scope is required' ], $config->validate( false ) );
	}


	#[Test]
	public function validateThrowsServiceExceptionWithCode400(): void {
		$config = new config();

		$this->expectException( serviceException::class );
		$this->expectExceptionCode( 400 );
		$config->validate();
	}


	#[Test]
	public function validateThrowExceptionMessageJoinsAllErrors(): void {
		$config = new config();

		try {
			$config->validate();
			$this->fail( 'Expected serviceException was not thrown' );
		}
		catch( serviceException $e ) {
			$this->assertSame( 'Client ID is required, Client secret is required, Tenant is required', $e->getMessage() );
		}
	}


	#[Test]
	#[DataProvider( 'individualRequiredFieldProvider' )]
	public function eachRequiredFieldIsValidatedIndividually( string $fieldToBlank, string $expectedError ): void {
		$config              = $this->validConfig();
		$config->$fieldToBlank = '';

		$this->assertSame( [ $expectedError ], $config->validate( false ) );
	}


	public static function individualRequiredFieldProvider(): array {
		return [
			'clientId'     => [ 'clientId', 'Client ID is required' ],
			'clientSecret' => [ 'clientSecret', 'Client secret is required' ],
			'tenant'       => [ 'tenant', 'Tenant is required' ],
			'scope'        => [ 'scope', 'Scope is required' ],
		];
	}


	#[Test]
	public function validateForMailThrowsWhenFromAddressMissing(): void {
		$config = $this->validConfig();

		$this->expectException( serviceException::class );
		$this->expectExceptionMessage( 'Default from email address is required' );
		$config->validateForMail();
	}


	#[Test]
	public function validateForMailPassesWhenFromAddressPresent(): void {
		$config              = $this->validConfig();
		$config->fromAddress = 'noreply@example.com';

		$config->validateForMail();
		$this->addToAssertionCount( 1 );
	}


	#[Test]
	public function validateForMailAggregatesBaseAndMailErrors(): void {
		$config = new config();

		try {
			$config->validateForMail();
			$this->fail( 'Expected serviceException was not thrown' );
		}
		catch( serviceException $e ) {
			$this->assertSame(
				'Client ID is required, Client secret is required, Tenant is required, Default from email address is required',
				$e->getMessage()
			);
		}
	}


	#[Test]
	public function validateForFilesThrowsWhenDriveIdMissing(): void {
		$config = $this->validConfig();

		$this->expectException( serviceException::class );
		$this->expectExceptionMessage( 'Drive id is required to manage files' );
		$config->validateForFiles();
	}


	#[Test]
	public function validateForFilesPassesWhenDriveIdPresent(): void {
		$config          = $this->validConfig();
		$config->driveId = 'drive-id';

		$config->validateForFiles();
		$this->addToAssertionCount( 1 );
	}

}
