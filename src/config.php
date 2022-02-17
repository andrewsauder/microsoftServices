<?php

namespace andrewsauder\microsoftServices;


use andrewsauder\microsoftServices\exceptions\serviceException;


class config {

	public bool   $onBehalfOfFlow = true;

	public string $clientId       = "";

	public string $clientSecret   = "";

	public string $tenant         = "";

	public string $driveId        = "";

	public string $fromAddress    = "";

	public string $scope          = 'openid profile email offline_access';


	public function __construct() {
	}


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function validate( bool $throwServiceException=true ) : array {
		 $errors = [];
		if( empty($this->clientId)) {
			$errors[] = 'Client ID is required';
		}
		if( empty($this->clientSecret)) {
			$errors[] = 'Client secret is required';
		}
		if( empty($this->tenant)) {
			$errors[] = 'Tenant is required';
		}
		if( empty($this->scope)) {
			$errors[] = 'Scope is required';
		}

		if ( $throwServiceException && count( $errors)>0) {
			throw new serviceException( implode(', ', $errors), 400 );
		}

		return $errors;
	}

	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function validateForMail() {
		$errors = $this->validate( false );

		if( empty($this->fromAddress)) {
			$errors[] = 'Default from email address is required';
		}
		if ( count( $errors)>0) {
			throw new serviceException( implode(', ', $errors), 400 );
		}
	}

	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function validateForFiles() {
		$errors = $this->validate( false );

		if( empty($this->driveId)) {
			$errors[] = 'Drive id is required to manage files';
		}

		if ( count( $errors)>0) {
			throw new serviceException( implode(', ', $errors), 400 );
		}
	}

}