<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests;

use andrewsauder\microsoftServices\files;
use PHPUnit\Framework\Attributes\CoversClass;
use PHPUnit\Framework\Attributes\CoversMethod;
use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\Attributes\Test;
use PHPUnit\Framework\TestCase;

#[CoversClass( files::class )]
#[CoversMethod( files::class, 'sanitizeFileFolderName' )]
final class filesSanitizeTest extends TestCase {

	#[Test]
	#[DataProvider( 'sanitizeProvider' )]
	public function sanitizesNameAsExpected( string $input, string $expected ): void {
		$this->assertSame( $expected, files::sanitizeFileFolderName( $input ) );
	}


	public static function sanitizeProvider(): array {
		return [
			'clean name is untouched'              => [ 'Building 1', 'Building 1' ],
			'clean name with period is untouched'  => [ 'report.v2.pdf', 'report.v2.pdf' ],
			'removes every forbidden character'    => [ 'a~b#c%d&e*f{g}h\\i:j<k>l?m/n+o|p"q', 'abcdefghijklmnopq' ],
			'removes a single forbidden character' => [ 'in/valid', 'invalid' ],
			'collapses consecutive periods'        => [ 'a...b', 'a.b' ],
			'collapses long period runs'           => [ 'file....txt', 'file.txt' ],
			'trims leading and trailing periods'   => [ '.hidden.', 'hidden' ],
			'trims surrounding whitespace'         => [ '  hello  ', 'hello' ],
			'trims mixed periods and spaces'       => [ ' . spaced . ', 'spaced' ],
			'removes .files suffix'                => [ 'MyDocument.files', 'MyDocument' ],
			'removes _files suffix'                => [ 'photos_files', 'photos' ],
			'removes suffix case-insensitively'    => [ 'REPORT_FILES', 'REPORT' ],
			'removes german -Dateien suffix'       => [ 'Bilder-Dateien', 'Bilder' ],
			'removes only the first matched suffix'=> [ 'archive_files', 'archive' ],
			'applies every rule together'          => [ '  My..File/Name.files  ', 'My.FileName' ],
		];
	}


	#[Test]
	public function emptyInputProducesRandomHexName(): void {
		$this->assertMatchesRegularExpression( '/^[0-9a-f]{20}$/', files::sanitizeFileFolderName( '' ) );
	}


	#[Test]
	public function nameConsistingOnlyOfForbiddenCharactersProducesRandomHexName(): void {
		$this->assertMatchesRegularExpression( '/^[0-9a-f]{20}$/', files::sanitizeFileFolderName( '///\\\\:::' ) );
	}


	#[Test]
	public function nameThatIsEntirelyADisallowedSuffixProducesRandomHexName(): void {
		$this->assertMatchesRegularExpression( '/^[0-9a-f]{20}$/', files::sanitizeFileFolderName( '_files' ) );
	}


	#[Test]
	public function randomFallbackNamesAreNotStable(): void {
		//two independent empty inputs should each get their own random name
		$this->assertNotSame(
			files::sanitizeFileFolderName( '' ),
			files::sanitizeFileFolderName( '' )
		);
	}

}
