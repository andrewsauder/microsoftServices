<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests\components;

use andrewsauder\microsoftServices\components\envelope;
use PHPUnit\Framework\Attributes\CoversClass;
use PHPUnit\Framework\Attributes\Test;
use PHPUnit\Framework\TestCase;

#[CoversClass( envelope::class )]
final class envelopeTest extends TestCase {

	#[Test]
	public function defaultsRepresentASuccessfulEmptyResponse(): void {
		$envelope = new envelope();

		$this->assertSame( 200, $envelope->status );
		$this->assertFalse( $envelope->error );
		$this->assertSame( '', $envelope->message );
		$this->assertSame( [], $envelope->data );
	}


	#[Test]
	public function constructorAssignsAllProvidedValues(): void {
		$envelope = new envelope( 404, true, 'Not found', [ 'id' => 7 ] );

		$this->assertSame( 404, $envelope->status );
		$this->assertTrue( $envelope->error );
		$this->assertSame( 'Not found', $envelope->message );
		$this->assertSame( [ 'id' => 7 ], $envelope->data );
	}

}
