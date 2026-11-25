<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests\live;

use andrewsauder\microsoftServices\calendar;
use andrewsauder\microsoftServices\config;
use andrewsauder\microsoftServices\files;
use andrewsauder\microsoftServices\mail;
use andrewsauder\microsoftServices\user;
use PHPUnit\Framework\Attributes\CoversNothing;
use PHPUnit\Framework\Attributes\Group;
use PHPUnit\Framework\Attributes\Test;
use PHPUnit\Framework\TestCase;

/**
 * Calls a real Microsoft 365 tenant with an application token. The default test run excludes this group. Run it
 * against a test tenant only, because it sends mail and writes files:
 *
 *   MS_TENANT=... MS_CLIENT_ID=... MS_CLIENT_SECRET=... MS_MAILBOX=... MS_DRIVE_ID=... vendor/bin/phpunit --group live
 *
 * The app needs the application permissions User.Read.All, Calendars.Read, Mail.ReadWrite, Mail.Send, and
 * Files.ReadWrite.All. A test whose environment variables are missing is skipped.
 */
#[Group( 'live' )]
#[CoversNothing]
final class liveSmokeTest extends TestCase {

	private string $directory;


	protected function setUp(): void {
		$this->directory = sys_get_temp_dir() . '/liveSmokeTest-' . bin2hex( random_bytes( 6 ) );
		mkdir( $this->directory );
	}


	protected function tearDown(): void {
		$items = new \RecursiveIteratorIterator( new \RecursiveDirectoryIterator( $this->directory, \FilesystemIterator::SKIP_DOTS ), \RecursiveIteratorIterator::CHILD_FIRST );
		foreach( $items as $item ) {
			$item->isDir() ? rmdir( $item->getPathname() ) : unlink( $item->getPathname() );
		}
		rmdir( $this->directory );
	}


	private static function env( string $name ): string {
		$value = getenv( $name );
		if( $value===false || $value==='' ) {
			self::markTestSkipped( $name . ' is not set' );
		}

		return $value;
	}


	private static function config(): config {
		$config               = new config();
		$config->tenant       = self::env( 'MS_TENANT' );
		$config->clientId     = self::env( 'MS_CLIENT_ID' );
		$config->clientSecret = self::env( 'MS_CLIENT_SECRET' );
		$config->fromAddress  = self::env( 'MS_MAILBOX' );
		$config->driveId      = getenv( 'MS_DRIVE_ID' ) ?: '';

		return $config;
	}


	private function file( string $name, int $size ): string {
		$path = $this->directory . '/' . $name;
		file_put_contents( $path, random_bytes( $size ) );

		return $path;
	}


	#[Test]
	public function usersAndCalendars(): void {
		$config = self::config();

		$users = ( new user( $config ) )->allUsersInOrganization();
		$this->assertNotEmpty( $users );

		$calendarService = new calendar( $config );
		$calendars       = $calendarService->getCalendars( $config->fromAddress );
		$this->assertNotEmpty( $calendars );

		$now = new \DateTimeImmutable();
		$calendarService->getCalendarEvents( $config->fromAddress, (string)$calendars[ 0 ]->getId(), $now, $now->modify( '+30 days' ) );
		$this->assertLessThanOrEqual( 3, count( $calendarService->getNCalendarEventsOccurringDuringOrAfter( $config->fromAddress, (string)$calendars[ 0 ]->getId(), $now, 3 ) ) );
	}


	#[Test]
	public function draftWithALargeAttachmentThenReplyIntoTheThread(): void {
		$config = self::config();
		$mail   = new mail( $config );

		$this->assertNotEmpty( $mail->getFolders( $config->fromAddress, true ) );

		$mail->addAttachment( $this->file( 'large.bin', 5 * 1024 * 1024 ) );
		$draft = $mail->createDraft( $config->fromAddress, 'microsoftServices live smoke test', '<p>Draft with a 5 MB attachment</p>', $config->fromAddress );
		$mail->sendDraft( (string)$draft->getId(), $config->fromAddress );

		$reply = ( new mail( $config ) )->replyIntoThread( $config->fromAddress, (string)$draft->getId(), $config->fromAddress, [], '<p>Reply</p>' );
		$this->assertNotNull( $reply->getConversationId() );
	}


	#[Test]
	public function uploadListDownloadRenameMoveAndDeleteAFile(): void {
		self::env( 'MS_DRIVE_ID' );
		$files  = new files( self::config() );
		$folder = 'microsoftServices smoke ' . bin2hex( random_bytes( 4 ) );

		$parent = $files->createFolder( $folder );
		try {
			$path   = $this->file( 'big file.bin', 10 * 1024 * 1024 );
			$upload = $files->upload( $path, 'big file.bin', [ $folder ] );
			$this->assertSame( [], $upload->errors );
			$item = $upload->files[ 0 ];

			$this->assertSame( [ 'big file.bin' ], array_map( static fn( $driveItem ) => $driveItem->getName(), $files->list( [ $folder ] ) ) );

			$download = $files->downloadDriveItemById( (string)$item->getId(), $this->directory . '/download' );
			$this->assertSame( filesize( $path ), filesize( $download ) );

			$files->renameItem( (string)$item->getId(), 'renamed.bin' );
			$subfolder = $files->createFolder( 'sub', [ $folder ] );
			$moved     = $files->moveItem( (string)$item->getId(), (string)$subfolder->getId() );
			$this->assertSame( 'renamed.bin', $moved->getName() );
		}
		finally {
			$files->delete( (string)$parent->getId() );
		}
	}

}
