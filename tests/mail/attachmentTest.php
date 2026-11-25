<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests\mail;

use andrewsauder\microsoftServices\exceptions\serviceException;
use andrewsauder\microsoftServices\mail\attachment;
use PHPUnit\Framework\Attributes\CoversClass;
use PHPUnit\Framework\Attributes\Test;
use PHPUnit\Framework\TestCase;

#[CoversClass( attachment::class )]
final class attachmentTest extends TestCase {

	private string $directory;


	protected function setUp(): void {
		$this->directory = sys_get_temp_dir() . '/attachmentTest-' . bin2hex( random_bytes( 6 ) );
		mkdir( $this->directory );
	}


	protected function tearDown(): void {
		foreach( glob( $this->directory . '/*' ) ?: [] as $file ) {
			unlink( $file );
		}
		rmdir( $this->directory );
	}


	#[Test]
	public function fromFileReadsTheNameTypeAndSize(): void {
		$path = $this->directory . '/notes.txt';
		file_put_contents( $path, 'hello world' );

		$attachment = attachment::fromFile( $path );

		$this->assertSame( $path, $attachment->path );
		$this->assertSame( 'notes.txt', $attachment->name );
		$this->assertSame( 'text/plain', $attachment->contentType );
		$this->assertSame( 11, $attachment->size );
	}


	#[Test]
	public function fromFileRejectsAMissingFile(): void {
		$this->expectException( serviceException::class );
		$this->expectExceptionCode( 500 );

		attachment::fromFile( $this->directory . '/missing.pdf' );
	}


	#[Test]
	public function fromFileRejectsAFileLargerThanAnUploadSessionAccepts(): void {
		$path   = $this->directory . '/huge.bin';
		$handle = fopen( $path, 'wb' );
		ftruncate( $handle, attachment::UPLOAD_SESSION_MAX_BYTES + 1 );
		fclose( $handle );

		$this->expectException( serviceException::class );
		$this->expectExceptionCode( 400 );

		attachment::fromFile( $path );
	}


	#[Test]
	public function aFileUpToTheInlineLimitIsPostedAndALargerOneUsesAnUploadSession(): void {
		$this->assertFalse( ( new attachment( '/x', 'x', 'application/pdf', 1 ) )->usesUploadSession() );
		$this->assertFalse( ( new attachment( '/x', 'x', 'application/pdf', attachment::INLINE_MAX_BYTES ) )->usesUploadSession() );
		$this->assertTrue( ( new attachment( '/x', 'x', 'application/pdf', attachment::INLINE_MAX_BYTES + 1 ) )->usesUploadSession() );
	}


	#[Test]
	public function anInlineFileStaysUnderTheGraphRequestLimitOnceEncoded(): void {
		$this->assertLessThan( 4 * 1024 * 1024, (int)ceil( attachment::INLINE_MAX_BYTES / 3 ) * 4 );
	}


	#[Test]
	public function eachUploadRangeStaysUnderFourMegabytes(): void {
		$this->assertLessThan( 4 * 1024 * 1024, attachment::CHUNK_BYTES );
	}


	#[Test]
	public function fileAttachmentBodyCarriesTheBase64Content(): void {
		$path = $this->directory . '/report.pdf';
		file_put_contents( $path, 'pdf bytes' );

		$body = attachment::fromFile( $path )->fileAttachmentBody();

		$this->assertSame( '#microsoft.graph.fileAttachment', $body[ '@odata.type' ] );
		$this->assertSame( 'report.pdf', $body[ 'name' ] );
		$this->assertSame( base64_encode( 'pdf bytes' ), $body[ 'contentBytes' ] );
	}


	#[Test]
	public function uploadSessionBodyDescribesTheFile(): void {
		$body = ( new attachment( '/tmp/scan.pdf', 'scan.pdf', 'application/pdf', 5_000_000 ) )->uploadSessionBody();

		$this->assertSame( [
			'AttachmentItem' => [
				'attachmentType' => 'file',
				'name'           => 'scan.pdf',
				'size'           => 5_000_000,
				'contentType'    => 'application/pdf'
			]
		], $body );
	}


	#[Test]
	public function byteRangesCoverTheFileInOrderWithoutGapsOrOverlap(): void {
		$this->assertSame( [ [ 0, 3 ], [ 4, 7 ], [ 8, 9 ] ], attachment::byteRanges( 10, 4 ) );
		$this->assertSame( [ [ 0, 3 ], [ 4, 7 ] ], attachment::byteRanges( 8, 4 ) );
		$this->assertSame( [ [ 0, 0 ] ], attachment::byteRanges( 1, 4 ) );
		$this->assertSame( [], attachment::byteRanges( 0, 4 ) );
	}


	#[Test]
	public function byteRangesForARealFileEndOnItsLastByte(): void {
		$size   = 7_500_000;
		$ranges = attachment::byteRanges( $size );

		$this->assertSame( 0, $ranges[ 0 ][ 0 ] );
		$this->assertSame( $size - 1, $ranges[ count( $ranges ) - 1 ][ 1 ] );
		$this->assertCount( (int)ceil( $size / attachment::CHUNK_BYTES ), $ranges );
	}


	#[Test]
	public function contentRangeUsesTheInclusiveGraphFormat(): void {
		$this->assertSame( 'bytes 0-3276799/7500000', attachment::contentRange( 0, 3_276_799, 7_500_000 ) );
	}

}
