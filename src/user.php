<?php

namespace andrewsauder\microsoftServices;

use andrewsauder\microsoftServices\exceptions\serviceException;
use GuzzleHttp\Exception\GuzzleException;
use Microsoft\Graph\Exception\GraphException;

class user extends \andrewsauder\microsoftServices\components\service {

	private array $attachments = [];


	/**
	 * @param string|null $userAccessToken    Provide user token. If config.onBehalfOfFlow is enabled, the provided token will be exchanged for an access token for this API. If config.onBehalfOfFlow is not enabled, this token will be used
	 *                                        for the request. If no token is provided, an application token will be generated
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function __construct( \andrewsauder\microsoftServices\config $config, ?string $userAccessToken = null ) {
		parent::__construct( $config, $userAccessToken );
	}


	/**
	 * @return \Microsoft\Graph\Model\User[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function allUsersInOrganization( ?string $nextLink=null ): array {
		try {
			if($nextLink===null) {
				$nextLink = '/users';
			}
			$graphResponse = $this->graph->createRequest( 'GET', $nextLink  )->execute();
			if($graphResponse->getNextLink()) {
				error_log($graphResponse->getNextLink());
				return array_merge( $graphResponse->getResponseAsObject( \Microsoft\Graph\Model\User::class ), $this->allUsersInOrganization( $graphResponse->getNextLink() ) );
			}
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to get user', $e->getCode(), $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to get user', $e->getCode(), $e );
		}

		return $graphResponse->getResponseAsObject( \Microsoft\Graph\Model\User::class );
	}


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getUserByEmail( string $email ): \Microsoft\Graph\Model\User {
		try {
			//get application access token
			$accessToken = $this->getMicrosoftAccessToken();

			$graph = new \Microsoft\Graph\Graph();
			$graph->setAccessToken( $accessToken );

			$response = $graph->createRequest( 'GET', '/users/' . $email )->setReturnType( \Microsoft\Graph\Model\User::class )->execute();
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to get user', $e->getCode(), $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to get user', $e->getCode(), $e );
		}
		return $response;
	}

}