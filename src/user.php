<?php

namespace andrewsauder\microsoftServices;

use Microsoft\Graph\Generated\Models\User as GraphUser;
use Microsoft\Graph\Generated\Users\UsersRequestBuilderGetRequestConfiguration;

class user extends \andrewsauder\microsoftServices\components\service {


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
	 * Every user in the organization, from every page of the /users collection
	 *
	 * @return \Microsoft\Graph\Generated\Models\User[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function allUsersInOrganization(): array {
		try {
			return $this->collectAll( $this->graph->users()->get()->wait() );
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to get users', $e );
		}
	}


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getUserByUserPrincipalName( string $upn ): GraphUser {
		try {
			$user = $this->graph->users()->byUserId( $upn )->get()->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to get user', $e );
		}

		if( !$user instanceof GraphUser ) {
			throw new \andrewsauder\microsoftServices\exceptions\serviceException( 'Failed to get user', 404 );
		}

		return $user;
	}


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getUserById( string $externalId ): GraphUser {
		return $this->getUserByUserPrincipalName( $externalId );
	}


	/**
	 * @param string $filter The filter command to pass into the MS Graph $filter url variable.
	 *                       Ex: to run Graph command /users?$filter=startswith(userPrincipalName,'asauder')
	 *                       pass just <i>startswith(userPrincipalName,'asauder')</i> to this function param
	 *
	 * @return \Microsoft\Graph\Generated\Models\User[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 * @see https://learn.microsoft.com/en-us/graph/filter-query-parameter?tabs=http
	 */
	public function getUsersByFilter( string $filter ): array {
		$requestConfiguration = new UsersRequestBuilderGetRequestConfiguration(
			queryParameters: UsersRequestBuilderGetRequestConfiguration::createQueryParameters( filter: $filter )
		);

		try {
			return $this->collectAll( $this->graph->users()->get( $requestConfiguration )->wait() );
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to get users', $e );
		}
	}

}
