<?php

namespace andrewsauder\microsoftServices\components;

class upload {

	/**
	 * @OA\Property()
	 * @var \Microsoft\Graph\Model\DriveItem[]
	 */
	public array                          $files        = [];

	/**
	 * @OA\Property()
	 * @var \andrewsauder\microsoftServices\components\envelope[]
	 */
	public array                          $errors     = [];


	public function __construct() {

	}

	public function merge( upload $upload ) {
		$this->files = array_merge( $this->files, $upload->files );
		$this->errors = array_merge( $this->errors, $upload->errors );
	}

}
