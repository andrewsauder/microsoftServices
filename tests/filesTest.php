<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests;

use andrewsauder\microsoftServices\components\envelope;
use andrewsauder\microsoftServices\exceptions\serviceException;
use andrewsauder\microsoftServices\files;
use andrewsauder\microsoftServices\mail\attachment;
use andrewsauder\microsoftServices\tests\support\graphTestCase;
use andrewsauder\microsoftServices\tests\support\testFiles;
use Microsoft\Graph\Generated\Models\DriveItem;
use PHPUnit\Framework\Attributes\CoversClass;
use PHPUnit\Framework\Attributes\Test;

#[CoversClass( files::class )]
final class filesTest extends graphTestCase {

	private const DRIVE = self::GRAPH . '/drives/drive-1';

	private string $directory;


	protected function setUp(): void {
		parent::setUp();
		$this->directory = sys_get_temp_dir() . '/filesTest-' . bin2hex( random_bytes( 6 ) );
		mkdir( $this->directory );
	}


	protected function tearDown(): void {
		$items = new \RecursiveIteratorIterator( new \RecursiveDirectoryIterator( $this->directory, \FilesystemIterator::SKIP_DOTS ), \RecursiveIteratorIterator::CHILD_FIRST );
		foreach( $items as $item ) {
			$item->isDir() ? rmdir( $item->getPathname() ) : unlink( $item->getPathname() );
		}
		rmdir( $this->directory );
		parent::tearDown();
	}


	private function files(): testFiles {
		return new testFiles( self::config(), self::TOKEN, 'Projects' );
	}


	private function file( string $name, int $size ): string {
		$path = $this->directory . '/' . $name;
		file_put_contents( $path, str_repeat( 'a', $size ) );

		return $path;
	}


	#[Test]
	public function listEncodesEachPathSegmentAndRecursesIntoFolders(): void {
		$this->queueJson( [ 'value' => [
			[ 'id' => 'folder-1', 'name' => 'Sub folder', 'folder' => [ 'childCount' => 1 ] ],
			[ 'id' => 'file-1', 'name' => 'a.txt', 'file' => [ 'mimeType' => 'text/plain' ] ],
		] ] );
		$this->queueJson( [ 'value' => [ [ 'id' => 'file-2', 'name' => 'b.txt' ] ] ] );

		$items = $this->files()->list( [ '2021-0001', 'Building 1' ] );

		$this->assertRequestCount( 2 );
		$this->assertSame( self::DRIVE . '/root:/Projects/2021-0001/Building%201:/children', $this->path( 0 ) );
		$this->assertSame( self::DRIVE . '/root:/Projects/2021-0001/Building%201/Sub%20folder:/children', $this->path( 1 ) );
		$this->assertSame( 'Bearer ' . self::TOKEN, $this->request( 0 )->getHeaderLine( 'Authorization' ) );

		$this->assertCount( 2, $items );
		$this->assertSame( 'b.txt', $items[ 0 ]->getChildren()[ 0 ]->getName() );
		$this->assertNull( $items[ 1 ]->getChildren() );
	}


	#[Test]
	public function listFollowsEveryNextLink(): void {
		$this->queueJson( [
			'value'           => [ [ 'id' => 'file-1', 'name' => 'a.txt' ] ],
			'@odata.nextLink' => self::DRIVE . '/items/root/children?$skiptoken=2'
		] );
		$this->queueJson( [ 'value' => [ [ 'id' => 'file-2', 'name' => 'b.txt' ] ] ] );

		$items = $this->files()->listById( 'root' );

		$this->assertSame( self::DRIVE . '/items/root/children', $this->path( 0 ) );
		$this->assertSame( [ 'a.txt', 'b.txt' ], array_map( static fn( DriveItem $item ) => $item->getName(), $items ) );
	}


	#[Test]
	public function listOfAFolderThatDoesNotExistIsEmpty(): void {
		$this->queueError( 404, 'itemNotFound', 'The resource could not be found.' );
		$this->expectErrorLog();

		$this->assertSame( [], $this->files()->list( [ 'missing' ] ) );
	}


	#[Test]
	public function getDriveItemOfAnEmptyPathIsTheDriveRoot(): void {
		$this->queueJson( [ 'id' => 'root-id', 'name' => 'root' ] );

		$item = $this->files()->getDriveItem( [] );

		$this->assertSame( self::DRIVE . '/root', $this->path( 0 ) );
		$this->assertSame( 'root-id', $item->getId() );
	}


	#[Test]
	public function uploadPutsASmallFileInOneRequest(): void {
		$this->queueJson( [ 'id' => 'file-1', 'name' => 'test.txt' ], 201 );

		$upload = $this->files()->upload( $this->file( 'test.txt', 12 ), 'test.txt', [ 'a b' ], 'rename' );

		$this->assertSame( [], $upload->errors );
		$this->assertSame( 'file-1', $upload->files[ 0 ]->getId() );
		$this->assertSame( 'PUT', $this->request( 0 )->getMethod() );
		$this->assertSame( self::DRIVE . '/root:/Projects/a%20b/test.txt:/content', $this->path( 0 ) );
		$this->assertSame( [ '@microsoft.graph.conflictBehavior' => 'rename' ], $this->query( 0 ) );
		$this->assertSame( str_repeat( 'a', 12 ), (string)$this->request( 0 )->getBody() );
	}


	#[Test]
	public function uploadSendsALargeFileThroughAnUploadSession(): void {
		$size = files::SIMPLE_UPLOAD_MAX_BYTES + 1;

		$this->queueJson( [ 'uploadUrl' => 'https://contoso.sharepoint.com/upload/1', 'nextExpectedRanges' => [ '0-' ] ] );
		$this->queueJson( [ 'nextExpectedRanges' => [ attachment::CHUNK_BYTES . '-' ] ], 202 );
		$this->queueJson( [ 'id' => 'file-1', 'name' => 'big.bin' ], 201 );
		$this->queueJson( [ 'id' => 'file-1', 'name' => 'big.bin' ] );

		$upload = $this->files()->upload( $this->file( 'big.bin', $size ), 'big.bin', [ 'a' ] );

		$this->assertSame( [], $upload->errors );
		$this->assertRequestCount( 4 );

		$this->assertSame( self::DRIVE . '/root:/Projects/a/big.bin:/createUploadSession', $this->path( 0 ) );
		$this->assertSame( 'replace', $this->jsonBody( 0 )[ 'item' ][ '@microsoft.graph.conflictBehavior' ] );
		$this->assertSame( 'big.bin', $this->jsonBody( 0 )[ 'item' ][ 'name' ] );

		$this->assertSame( 'https://contoso.sharepoint.com/upload/1', (string)$this->request( 1 )->getUri() );
		$this->assertFalse( $this->request( 1 )->hasHeader( 'Authorization' ), 'The upload URL carries its own token' );
		$this->assertSame( 'bytes ' . attachment::CHUNK_BYTES . '-' . ( $size - 1 ) . '/' . $size, $this->request( 2 )->getHeaderLine( 'Content-Range' ) );

		$this->assertSame( 'GET', $this->request( 3 )->getMethod() );
		$this->assertSame( self::DRIVE . '/root:/Projects/a/big.bin:', $this->path( 3 ) );
		$this->assertSame( 'file-1', $upload->files[ 0 ]->getId() );
	}


	#[Test]
	public function aFailedUploadIsAnErrorInTheResponse(): void {
		$this->queueError( 409, 'nameAlreadyExists', 'The name is already in use.' );
		$this->expectErrorLog();

		$upload = $this->files()->upload( $this->file( 'test.txt', 1 ), 'test.txt', [], 'fail' );

		$this->assertSame( [], $upload->files );
		$this->assertInstanceOf( envelope::class, $upload->errors[ 0 ] );
		$this->assertSame( 409, $upload->errors[ 0 ]->status );
		$this->assertSame( 'test.txt did not upload: The name is already in use.', $upload->errors[ 0 ]->message );
	}


	#[Test]
	public function uploadBase64EncodedContentUploadsTheDecodedBytes(): void {
		$this->queueJson( [ 'id' => 'file-1', 'name' => 'hello.txt' ], 201 );

		$upload = $this->files()->uploadBase64EncodedContent( base64_encode( 'hello' ), 'text/plain', 'hello.txt' );

		$this->assertSame( 'file-1', $upload->files[ 0 ]->getId() );
		$this->assertSame( 'hello', (string)$this->request( 0 )->getBody() );
	}


	#[Test]
	public function updateSendsTheNameParentAndOtherKeys(): void {
		$this->queueJson( [ 'id' => 'file-1', 'name' => 'new.txt' ] );

		$item = $this->files()->update( 'file-1', [
			'name'                              => 'new.txt',
			'parentReference'                   => [ 'id' => 'folder-2' ],
			'@microsoft.graph.conflictBehavior' => 'rename',
		] );

		$this->assertSame( 'PATCH', $this->request( 0 )->getMethod() );
		$this->assertSame( self::DRIVE . '/items/file-1', $this->path( 0 ) );
		$body = $this->jsonBody( 0 );
		$this->assertSame( 'new.txt', $body[ 'name' ] );
		$this->assertSame( 'folder-2', $body[ 'parentReference' ][ 'id' ] );
		$this->assertSame( 'rename', $body[ '@microsoft.graph.conflictBehavior' ] );
		$this->assertSame( 'new.txt', $item->getName() );
	}


	#[Test]
	public function updateWithNoChangesSendsNoPatch(): void {
		$this->queueJson( [ 'id' => 'file-1' ] );

		$this->files()->update( 'file-1', [] );

		$this->assertSame( 'GET', $this->request( 0 )->getMethod() );
	}


	#[Test]
	public function downloadDriveItemByIdWritesTheContentToTheFile(): void {
		$this->queueJson( [ 'id' => 'file-1', 'name' => 'report.txt' ] );
		$this->mockHandler->append( new \GuzzleHttp\Psr7\Response( 200, [ 'Content-Type' => 'application/octet-stream' ], 'file content' ) );

		$filePathName = $this->files()->downloadDriveItemById( 'file-1', $this->directory . '/downloads' );

		$this->assertSame( $this->directory . '/downloads/file-1/report.txt', $filePathName );
		$this->assertSame( 'file content', file_get_contents( $filePathName ) );
		$this->assertSame( 'id,name', $this->query( 0 )[ '$select' ] );
		$this->assertSame( self::DRIVE . '/items/file-1/content', $this->path( 1 ) );
	}


	#[Test]
	public function createFolderAsksGraphToFailOnANameConflict(): void {
		$this->queueJson( [ 'id' => 'folder-1', 'name' => 'New', 'folder' => [ 'childCount' => 0 ] ], 201 );

		$folder = $this->files()->createFolder( 'New', [ '2021-0001' ] );

		$this->assertSame( self::DRIVE . '/root:/Projects/2021-0001:/children', $this->path( 0 ) );
		$body = $this->jsonBody( 0 );
		$this->assertSame( 'New', $body[ 'name' ] );
		$this->assertSame( [], $body[ 'folder' ] );
		$this->assertSame( 'fail', $body[ '@microsoft.graph.conflictBehavior' ] );
		$this->assertSame( 'folder-1', $folder->getId() );
	}


	#[Test]
	public function getDriveItemThumbnailByIdPicksTheSize(): void {
		$this->queueJson( [ 'id' => '0', 'medium' => [ 'url' => 'https://example.test/medium.jpg', 'width' => 176 ] ] );

		$thumbnail = $this->files()->getDriveItemThumbnailById( 'file-1' );

		$this->assertSame( self::DRIVE . '/items/file-1/thumbnails/0', $this->path( 0 ) );
		$this->assertSame( 'https://example.test/medium.jpg', $thumbnail->getUrl() );
	}


	#[Test]
	public function getDriveItemThumbnailByIdRejectsAnUnknownSize(): void {
		$this->expectException( serviceException::class );
		$this->expectExceptionCode( 400 );

		$this->files()->getDriveItemThumbnailById( 'file-1', 0, 'c300x400' );
	}


	#[Test]
	public function getDriveItemPreviewByIdSendsThePageAndZoom(): void {
		$this->queueJson( [ 'getUrl' => 'https://example.test/preview' ] );

		$preview = $this->files()->getDriveItemPreviewById( 'file-1', 3, 2 );

		$this->assertSame( self::DRIVE . '/items/file-1/preview', $this->path( 0 ) );
		$this->assertEquals( [ 'page' => '3', 'zoom' => 2 ], $this->jsonBody( 0 ) );
		$this->assertSame( 'https://example.test/preview', $preview->getGetUrl() );
	}


	#[Test]
	public function deleteSendsADelete(): void {
		$this->queueEmpty( 204 );

		$this->files()->delete( 'file-1' );

		$this->assertSame( 'DELETE', $this->request( 0 )->getMethod() );
		$this->assertSame( self::DRIVE . '/items/file-1', $this->path( 0 ) );
	}

}
