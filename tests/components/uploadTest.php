<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests\components;

use andrewsauder\microsoftServices\components\upload;
use PHPUnit\Framework\Attributes\CoversClass;
use PHPUnit\Framework\Attributes\Test;
use PHPUnit\Framework\TestCase;

#[CoversClass( upload::class )]
final class uploadTest extends TestCase {

	#[Test]
	public function newUploadStartsEmpty(): void {
		$upload = new upload();

		$this->assertSame( [], $upload->files );
		$this->assertSame( [], $upload->errors );
	}


	#[Test]
	public function mergeCombinesFilesAndErrorsFromBothUploads(): void {
		$first          = new upload();
		$first->files   = [ 'a', 'b' ];
		$first->errors  = [ 'e1' ];

		$second         = new upload();
		$second->files  = [ 'c' ];
		$second->errors = [ 'e2', 'e3' ];

		$first->merge( $second );

		$this->assertSame( [ 'a', 'b', 'c' ], $first->files );
		$this->assertSame( [ 'e1', 'e2', 'e3' ], $first->errors );
	}


	#[Test]
	public function mergeWithEmptyUploadLeavesTargetUnchanged(): void {
		$target        = new upload();
		$target->files = [ 'only' ];

		$target->merge( new upload() );

		$this->assertSame( [ 'only' ], $target->files );
		$this->assertSame( [], $target->errors );
	}


	#[Test]
	public function mergeDoesNotMutateTheSourceUpload(): void {
		$target        = new upload();
		$target->files = [ 'x' ];

		$source        = new upload();
		$source->files = [ 'y' ];

		$target->merge( $source );

		$this->assertSame( [ 'y' ], $source->files, 'the merged-in upload should be left untouched' );
	}

}
