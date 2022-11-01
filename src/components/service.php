<?php

namespace andrewsauder\microsoftServices\components;

use andrewsauder\microsoftServices\auth;

class service {

	protected \andrewsauder\microsoftServices\config $config;

	protected ?string $userAccessToken = null;

	protected \Microsoft\Graph\Graph $graph;


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function __construct( \andrewsauder\microsoftServices\config $config, ?string $userAccessToken = null ) {
		$this->config          = $config;
		$this->userAccessToken = $userAccessToken;

		$this->newGraph();
	}


	/**
	 * @return \Microsoft\Graph\Graph
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	private function newGraph(): \Microsoft\Graph\Graph {
		$accessToken = $this->getMicrosoftAccessToken();

		$this->graph = new \Microsoft\Graph\Graph();
		$this->graph->setAccessToken( $accessToken );

		return $this->graph;
	}

	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	private function graphRequestFollowPaging(): \Microsoft\Graph\Graph {
		$accessToken = $this->getMicrosoftAccessToken();

		$this->graph = new \Microsoft\Graph\Graph();
		$this->graph->setAccessToken( $accessToken );

		return $this->graph;
	}


	/**
	 * @return string
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	protected function getMicrosoftAccessToken(): string {
		$microsoftAuth = new auth( $this->config );

		if( !$this->config->onBehalfOfFlow && isset( $this->userAccessToken ) ) {
			return $this->userAccessToken;
		}
		//get OBO user access token
		if( $this->config->onBehalfOfFlow && isset( $this->userAccessToken ) ) {
			return (string)$microsoftAuth->getAccessToken( $this->userAccessToken );
		}
		//get access token
		else {
			return $microsoftAuth->getApplicationAccessToken();
		}
	}

}
