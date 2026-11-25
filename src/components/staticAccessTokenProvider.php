<?php

namespace andrewsauder\microsoftServices\components;

use Http\Promise\FulfilledPromise;
use Http\Promise\Promise;
use Microsoft\Kiota\Abstractions\Authentication\AccessTokenProvider;
use Microsoft\Kiota\Abstractions\Authentication\AllowedHostsValidator;

/**
 * Gives the Graph client an access token that the caller already holds. The token is sent as it is: this class
 * does not refresh it. It is sent to Microsoft Graph hosts only.
 */
final class staticAccessTokenProvider implements AccessTokenProvider {

	private AllowedHostsValidator $allowedHostsValidator;


	public function __construct( private readonly string $accessToken ) {
		$this->allowedHostsValidator = new AllowedHostsValidator( [ 'graph.microsoft.com' ] );
	}


	/**
	 * @param array<string, mixed> $additionalAuthenticationContext
	 *
	 * @return Promise<string|null>
	 */
	public function getAuthorizationTokenAsync( string $url, array $additionalAuthenticationContext = [] ): Promise {
		if( !$this->allowedHostsValidator->isUrlHostValid( $url ) ) {
			return new FulfilledPromise( null );
		}

		return new FulfilledPromise( $this->accessToken );
	}


	public function getAllowedHostsValidator(): AllowedHostsValidator {
		return $this->allowedHostsValidator;
	}

}
