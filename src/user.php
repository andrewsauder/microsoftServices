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
	public function allUsersInOrganization( ?string $nextLink = null ): array {
		try {
			if( $nextLink===null ) {
				$nextLink = '/users';
			}
			$graphResponse = $this->graph->createRequest( 'GET', $nextLink )->execute();
			if( $graphResponse->getNextLink() ) {
				error_log( $graphResponse->getNextLink() );
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
	public function getUserByUserPrincipalName( string $upn ): \Microsoft\Graph\Model\User {
		try {
			//get application access token
			$accessToken = $this->getMicrosoftAccessToken();

			$graph = new \Microsoft\Graph\Graph();
			$graph->setAccessToken( $accessToken );

			$response = $graph->createRequest( 'GET', '/users/' . $upn )->setReturnType( \Microsoft\Graph\Model\User::class )->execute();
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


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getUserByExternalId( string $externalId ): \Microsoft\Graph\Model\User {
		return $this->getUserByUserPrincipalName( $externalId );
	}


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 * @deprecated use getUserByUserPrincipalName instead. This function has always operate on UPN, not email. Preserved in v2 for backwards compatibility only. Will be removed in v3 or changed to actually search for user by mail attribute
	 */
	public function getUserByEmail( string $email ): \Microsoft\Graph\Model\User {
		return $this->getUserByUserPrincipalName( $email );
	}


	/**
	 * @param string $filter The filter command to pass into the MS Graph $filter url variable.
	 *                       Ex: to run Graph command /users?$filter=startswith(userPrincipalName,'asauder')
	 *                       pass just <i>startswith(userPrincipalName,'asauder')</i> to this function param
	 *
	 * @return \Microsoft\Graph\Model\User[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 * @see https://learn.microsoft.com/en-us/graph/filter-query-parameter?tabs=http
	 *
	 */
	public function getUsersByFilter( string $filter ): array {
		try {
			//get application access token
			$accessToken = $this->getMicrosoftAccessToken();

			$graph = new \Microsoft\Graph\Graph();
			$graph->setAccessToken( $accessToken );

			$response = $graph->createRequest( 'GET', '/users?$filter=' . $filter )->setReturnType( \Microsoft\Graph\Model\User::class )->execute();
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
