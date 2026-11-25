<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests;

use andrewsauder\microsoftServices\exceptions\serviceException;
use andrewsauder\microsoftServices\mail;
use andrewsauder\microsoftServices\mail\attachment;
use andrewsauder\microsoftServices\tests\support\graphTestCase;
use andrewsauder\microsoftServices\tests\support\testMail;
use Microsoft\Graph\Generated\Models\FileAttachment;
use PHPUnit\Framework\Attributes\CoversClass;
use PHPUnit\Framework\Attributes\Test;

#[CoversClass( mail::class )]
final class mailTest extends graphTestCase {

	private const MAILBOX = self::GRAPH . '/users/noreply%40contoso.example';

	private string $directory;


	protected function setUp(): void {
		parent::setUp();
		$this->directory = sys_get_temp_dir() . '/mailTest-' . bin2hex( random_bytes( 6 ) );
		mkdir( $this->directory );
	}


	protected function tearDown(): void {
		foreach( glob( $this->directory . '/*' ) ?: [] as $file ) {
			unlink( $file );
		}
		rmdir( $this->directory );
		parent::tearDown();
	}


	private function file( string $name, int $size ): string {
		$path = $this->directory . '/' . $name;
		file_put_contents( $path, str_repeat( 'a', $size ) );

		return $path;
	}


	#[Test]
	public function sendPostsOneSendMailRequestWithTheAttachments(): void {
		$this->queueEmpty( 202 );

		$mail = new testMail( self::config(), self::TOKEN );
		$mail->addAttachment( $this->file( 'notes.txt', 5 ) );
		$mail->send( [ 'ann@contoso.example', 'bo@contoso.example' ], 'Subject', '<p>Hi</p>' );

		$this->assertRequestCount( 1 );
		$this->assertSame( 'POST', $this->request( 0 )->getMethod() );
		$this->assertSame( self::MAILBOX . '/sendMail', $this->path( 0 ) );

		$message = $this->jsonBody( 0 )[ 'Message' ];
		$this->assertSame( 'Subject', $message[ 'subject' ] );
		$this->assertSame( [ 'content' => '<p>Hi</p>', 'contentType' => 'html' ], $message[ 'body' ] );
		$this->assertSame( 'noreply@contoso.example', $message[ 'from' ][ 'emailAddress' ][ 'address' ] );
		$this->assertSame( [ 'ann@contoso.example', 'bo@contoso.example' ], array_map( static fn( array $recipient ) => $recipient[ 'emailAddress' ][ 'address' ], $message[ 'toRecipients' ] ) );
		$this->assertSame( '#microsoft.graph.fileAttachment', $message[ 'attachments' ][ 0 ][ '@odata.type' ] );
		$this->assertSame( base64_encode( 'aaaaa' ), $message[ 'attachments' ][ 0 ][ 'contentBytes' ] );
	}


	#[Test]
	public function aRejectedSendThrowsWithTheHttpStatus(): void {
		$this->queueError( 413, 'RequestEntityTooLarge', 'The request is too large.' );
		$this->expectErrorLog();

		$this->expectException( serviceException::class );
		$this->expectExceptionCode( 413 );

		( new testMail( self::config(), self::TOKEN ) )->send( 'ann@contoso.example', 'Subject', 'Body' );
	}


	#[Test]
	public function createDraftAttachesASmallFileInOnePostAndALargeFileInRanges(): void {
		$largeSize = attachment::CHUNK_BYTES + 1000;

		$this->queueJson( [ 'id' => 'draft-1' ], 201 );
		$this->queueJson( [ 'id' => 'attachment-1' ], 201 );
		$this->queueJson( [ 'uploadUrl' => 'https://outlook.office.com/upload/draft-1', 'nextExpectedRanges' => [ '0-' ] ], 201 );
		$this->queueEmpty( 200 );
		$this->queueEmpty( 201 );

		$mail = new testMail( self::config(), self::TOKEN );
		$mail->addAttachment( $this->file( 'small.txt', 10 ) );
		$mail->addAttachment( $this->file( 'large.bin', $largeSize ) );
		$draft = $mail->createDraft( 'ann@contoso.example', 'Subject', 'Body', '', 'cc@contoso.example' );

		$this->assertSame( 'draft-1', $draft->getId() );
		$this->assertRequestCount( 5 );

		$this->assertSame( self::MAILBOX . '/messages', $this->path( 0 ) );
		$this->assertSame( 'IdType="ImmutableId"', $this->request( 0 )->getHeaderLine( 'Prefer' ) );
		$this->assertSame( 'cc@contoso.example', $this->jsonBody( 0 )[ 'ccRecipients' ][ 0 ][ 'emailAddress' ][ 'address' ] );

		$this->assertSame( self::MAILBOX . '/messages/draft-1/attachments', $this->path( 1 ) );
		$this->assertSame( 'small.txt', $this->jsonBody( 1 )[ 'name' ] );

		$this->assertSame( self::MAILBOX . '/messages/draft-1/attachments/createUploadSession', $this->path( 2 ) );
		$this->assertSame( $largeSize, $this->jsonBody( 2 )[ 'AttachmentItem' ][ 'size' ] );

		$this->assertSame( 'PUT', $this->request( 3 )->getMethod() );
		$this->assertSame( 'https://outlook.office.com/upload/draft-1', (string)$this->request( 3 )->getUri() );
		$this->assertFalse( $this->request( 3 )->hasHeader( 'Authorization' ), 'The upload URL carries its own token' );
		$this->assertSame( 'bytes 0-' . ( attachment::CHUNK_BYTES - 1 ) . '/' . $largeSize, $this->request( 3 )->getHeaderLine( 'Content-Range' ) );
		$this->assertSame( 'bytes ' . attachment::CHUNK_BYTES . '-' . ( $largeSize - 1 ) . '/' . $largeSize, $this->request( 4 )->getHeaderLine( 'Content-Range' ) );
	}


	#[Test]
	public function createDraftDeletesTheDraftWhenAnAttachmentFails(): void {
		$this->queueJson( [ 'id' => 'draft-1' ], 201 );
		$this->queueError( 413, 'RequestEntityTooLarge', 'The request is too large.' );
		$this->queueEmpty( 204 );
		$this->expectErrorLog();

		$mail = new testMail( self::config(), self::TOKEN );
		$mail->addAttachment( $this->file( 'small.txt', 10 ) );

		try {
			$mail->createDraft( 'ann@contoso.example', 'Subject', 'Body', '' );
			$this->fail( 'Expected a serviceException' );
		}
		catch( serviceException $e ) {
			$this->assertSame( 413, $e->getCode() );
		}

		$this->assertRequestCount( 3 );
		$this->assertSame( 'DELETE', $this->request( 2 )->getMethod() );
		$this->assertSame( self::MAILBOX . '/messages/draft-1', $this->path( 2 ) );
	}


	#[Test]
	public function replyIntoThreadCreatesUpdatesAndSendsTheReply(): void {
		$this->queueJson( [ 'id' => 'reply-1' ], 201 );
		$this->queueJson( [ 'id' => 'reply-1', 'conversationId' => 'conversation-1' ] );
		$this->queueEmpty( 202 );

		$sent = ( new testMail( self::config(), self::TOKEN ) )->replyIntoThread( '', 'anchor-1', 'ann@contoso.example', [], 'Reply body' );

		$this->assertRequestCount( 3 );
		$this->assertSame( [ 'POST', 'PATCH', 'POST' ], array_map( fn( int $i ) => $this->request( $i )->getMethod(), [ 0, 1, 2 ] ) );
		$this->assertSame( self::MAILBOX . '/messages/anchor-1/createReply', $this->path( 0 ) );
		$this->assertSame( self::MAILBOX . '/messages/reply-1', $this->path( 1 ) );
		$this->assertSame( self::MAILBOX . '/messages/reply-1/send', $this->path( 2 ) );
		$this->assertSame( 'IdType="ImmutableId"', $this->request( 0 )->getHeaderLine( 'Prefer' ) );
		$this->assertSame( 'IdType="ImmutableId"', $this->request( 1 )->getHeaderLine( 'Prefer' ) );
		$this->assertSame( 'Reply body', $this->jsonBody( 1 )[ 'body' ][ 'content' ] );
		$this->assertSame( [], $this->jsonBody( 1 )[ 'bccRecipients' ] );
		$this->assertSame( 'conversation-1', $sent->getConversationId() );
	}


	#[Test]
	public function getLatestMessageInConversationQueriesTheNewestMessage(): void {
		$this->queueJson( [ 'value' => [] ] );

		$latest = ( new testMail( self::config(), self::TOKEN ) )->getLatestMessageInConversation( '', "it's-1" );

		$this->assertNull( $latest );
		$query = $this->query( 0 );
		$this->assertSame( "conversationId eq 'it''s-1'", $query[ '$filter' ] );
		$this->assertSame( 'receivedDateTime desc', $query[ '$orderby' ] );
		$this->assertSame( 'id,conversationId', $query[ '$select' ] );
		$this->assertSame( '1', $query[ '$top' ] );
	}


	#[Test]
	public function getAllMessagesAsksForTheBodyTypeAndImmutableIds(): void {
		$this->queueJson( [ 'value' => [ [ 'id' => 'message-1' ] ] ] );

		$messages = ( new testMail( self::config(), self::TOKEN ) )->getAllMessages( 'ann@contoso.example', \andrewsauder\microsoftServices\mail\headers\prefer::TEXT );

		$this->assertCount( 1, $messages );
		$this->assertSame( 'outlook.body-content-type="text", IdType="ImmutableId"', $this->request( 0 )->getHeaderLine( 'Prefer' ) );
	}


	#[Test]
	public function getFoldersCanIncludeHiddenFolders(): void {
		$this->queueJson( [ 'value' => [ [ 'id' => 'folder-1', 'displayName' => 'Inbox' ] ] ] );

		$folders = ( new testMail( self::config(), self::TOKEN ) )->getFolders( 'ann@contoso.example', true );

		$this->assertSame( self::GRAPH . '/users/ann%40contoso.example/mailFolders', $this->path( 0 ) );
		$this->assertSame( [ 'includeHiddenFolders' => 'true' ], $this->query( 0 ) );
		$this->assertSame( 'Inbox', $folders[ 0 ]->getDisplayName() );
	}


	#[Test]
	public function getAttachmentsReturnsFileAttachmentsWithTheBase64Content(): void {
		$this->queueJson( [ 'value' => [ [ '@odata.type' => '#microsoft.graph.fileAttachment', 'name' => 'a.txt', 'contentBytes' => base64_encode( 'hello' ) ] ] ] );

		$attachments = ( new testMail( self::config(), self::TOKEN ) )->getAttachments( 'ann@contoso.example', 'message-1' );

		$this->assertInstanceOf( FileAttachment::class, $attachments[ 0 ] );
		$this->assertSame( base64_encode( 'hello' ), (string)$attachments[ 0 ]->getContentBytes() );
	}


	#[Test]
	public function moveMessageReturnsTheMovedMessage(): void {
		$this->queueJson( [ 'id' => 'message-1', 'parentFolderId' => 'archive-id' ], 201 );

		$moved = ( new testMail( self::config(), self::TOKEN ) )->moveMessage( 'ann@contoso.example', 'message-1', 'archive' );

		$this->assertSame( self::GRAPH . '/users/ann%40contoso.example/messages/message-1/move', $this->path( 0 ) );
		$this->assertSame( [ 'DestinationId' => 'archive' ], $this->jsonBody( 0 ) );
		$this->assertSame( 'archive-id', $moved->getParentFolderId() );
	}


	#[Test]
	public function deleteMessageSendsADelete(): void {
		$this->queueEmpty( 204 );

		( new testMail( self::config(), self::TOKEN ) )->deleteMessage( 'ann@contoso.example', 'message-1' );

		$this->assertSame( 'DELETE', $this->request( 0 )->getMethod() );
		$this->assertSame( self::GRAPH . '/users/ann%40contoso.example/messages/message-1', $this->path( 0 ) );
	}

}
