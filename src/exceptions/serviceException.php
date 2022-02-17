<?php
namespace andrewsauder\microsoftServices\exceptions;


use JetBrains\PhpStorm\Pure;


class serviceException
	extends
	\Exception {

	#[Pure]
	public function __construct( $message, $code = 0, \Exception $previous = null ) {
		// make sure everything is assigned properly
		parent::__construct( $message, $code, $previous );
	}

}