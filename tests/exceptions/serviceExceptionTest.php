<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests\exceptions;

use andrewsauder\microsoftServices\exceptions\serviceException;
use GuzzleHttp\Exception\TransferException;
use GuzzleHttp\Psr7\Request;
use PHPUnit\Framework\Attributes\CoversClass;
use PHPUnit\Framework\Attributes\Test;
use PHPUnit\Framework\TestCase;

#[CoversClass( serviceException::class )]
final class serviceExceptionTest extends TestCase {

	#[Test]
	public function isAnException(): void {
		$this->assertInstanceOf( \Exception::class, new serviceException( 'boom' ) );
	}


	#[Test]
	public function retainsMessageAndCode(): void {
		$exception = new serviceException( 'something failed', 418 );

		$this->assertSame( 'something failed', $exception->getMessage() );
		$this->assertSame( 418, $exception->getCode() );
	}


	#[Test]
	public function codeDefaultsToZero(): void {
		$this->assertSame( 0, ( new serviceException( 'no code given' ) )->getCode() );
	}


	#[Test]
	public function previousDefaultsToNull(): void {
		$this->assertNull( ( new serviceException( 'no previous' ) )->getPrevious() );
	}


	#[Test]
	public function retainsExceptionAsPrevious(): void {
		$previous  = new \RuntimeException( 'root cause' );
		$exception = new serviceException( 'wrapped', 500, $previous );

		$this->assertSame( $previous, $exception->getPrevious() );
	}


	#[Test]
	public function acceptsGuzzleExceptionAsPrevious(): void {
		//GuzzleException only guarantees Throwable, not Exception - this is the case the constructor signature fix
		//had to support
		$previous  = new TransferException( 'transfer blew up', new Request( 'GET', 'https://graph.microsoft.com/v1.0/me' ) );
		$exception = new serviceException( 'wrapped guzzle error', 502, $previous );

		$this->assertSame( $previous, $exception->getPrevious() );
	}


	#[Test]
	public function acceptsAnyThrowableAsPrevious(): void {
		//an Error is a Throwable but not an Exception; the ?\Throwable signature must accept it
		$previous  = new \TypeError( 'bad type' );
		$exception = new serviceException( 'wrapped error', 500, $previous );

		$this->assertSame( $previous, $exception->getPrevious() );
	}

}
