<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests\mail\headers;

use andrewsauder\microsoftServices\mail\headers\prefer;
use PHPUnit\Framework\Attributes\CoversClass;
use PHPUnit\Framework\Attributes\Test;
use PHPUnit\Framework\TestCase;

#[CoversClass( prefer::class )]
final class preferTest extends TestCase {

	#[Test]
	public function backingValuesMatchTheGraphContentTypes(): void {
		$this->assertSame( 'html', prefer::HTML->value );
		$this->assertSame( 'text', prefer::TEXT->value );
	}


	#[Test]
	public function fromResolvesBackingValues(): void {
		$this->assertSame( prefer::HTML, prefer::from( 'html' ) );
		$this->assertSame( prefer::TEXT, prefer::from( 'text' ) );
	}


	#[Test]
	public function tryFromReturnsNullForUnknownValue(): void {
		$this->assertNull( prefer::tryFrom( 'markdown' ) );
	}


	#[Test]
	public function exposesExactlyTwoCases(): void {
		$this->assertSame( [ prefer::HTML, prefer::TEXT ], prefer::cases() );
	}

}
