<?php

namespace andrewsauder\microsoftServices\components;


use andrewsauder\microsoftServices\auth;


class service {

	protected \andrewsauder\microsoftServices\config $config;

	protected ?string                                $userAccessToken = null;


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function __construct( \andrewsauder\microsoftServices\config $config, ?string $userAccessToken = null ) {
		$config->validateForFiles();
		$this->config          = $config;
		$this->userAccessToken = $userAccessToken;
	}


	/**
	 * @return string
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	protected function getMicrosoftAccessToken() : string {
		$microsoftAuth = new auth( $this->config );

		if( !$this->config->onBehalfOfFlow && isset( $this->userAccessToken ) ) {
			return $this->userAccessToken;
		}
		//get OBO user access token
		if( $this->config->onBehalfOfFlow && isset( $this->userAccessToken ) ) {
			return (string) $microsoftAuth->getAccessToken( $this->userAccessToken );
		}
		//get access token
		else {
			return $microsoftAuth->getApplicationAccessToken();
		}
	}

}
