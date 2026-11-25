<?php

namespace andrewsauder\microsoftServices;



class auth {

	private \andrewsauder\microsoftServices\config   $config;

	private \TheNetworg\OAuth2\Client\Provider\Azure $provider;


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function __construct( \andrewsauder\microsoftServices\config $config ) {
		$config->validate();
		$this->config   = $config;
		$this->provider = new \TheNetworg\OAuth2\Client\Provider\Azure( (array) $config );
	}


	/**
	 * @param  string  $suppliedToken
	 *
	 * @return \TheNetworg\OAuth2\Client\Token\AccessToken
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getAccessToken( string $suppliedToken ) : \TheNetworg\OAuth2\Client\Token\AccessToken {
		$this->provider->defaultEndPointVersion = \TheNetworg\OAuth2\Client\Provider\Azure::ENDPOINT_VERSION_2_0;
		$this->provider->scope                  = $this->config->scope;

		try {
			$token = $this->provider->getAccessToken( 'jwt_bearer', [
				'scope'               => $this->provider->scope,
				'assertion'           => $suppliedToken,
				'requested_token_use' => 'on_behalf_of',
			] );
		}
		catch( \Exception $e ) {
			throw new \andrewsauder\microsoftServices\exceptions\serviceException( 'Microsoft authentication failed', 400, $e );
		}

		//the Azure provider always mints a TheNetworg AccessToken (which carries the id token claims); guard the
		//interface-level return type so downstream calls to getIdTokenClaims() are type-safe
		if( !$token instanceof \TheNetworg\OAuth2\Client\Token\AccessToken ) {
			throw new \andrewsauder\microsoftServices\exceptions\serviceException( 'Unexpected access token type returned by identity provider', 500 );
		}

		return $token;
	}


	/**
	 * @return \andrewsauder\microsoftServices\components\tokenInformation
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function verify() : \andrewsauder\microsoftServices\components\tokenInformation {
		$suppliedToken = str_replace( 'Bearer ', '', $_SERVER[ 'HTTP_AUTHORIZATION' ] );
		$token         = $this->getAccessToken( $suppliedToken );

		$claims  = $token->getIdTokenClaims();
		$expires = $token->getExpires();

		$tokenInformation                     = new components\tokenInformation();
		$tokenInformation->iss                = $claims[ 'iss' ] ?? '';
		$tokenInformation->aud                = $claims[ 'aud' ] ?? '';
		$tokenInformation->oid                = $claims[ 'oid' ] ?? '';
		$tokenInformation->sub                = $claims[ 'sub' ] ?? '';
		$tokenInformation->appId              = $claims[ 'appid' ] ?? '';
		$tokenInformation->name               = $claims[ 'name' ] ?? '';
		$tokenInformation->familyName         = $claims[ 'family_name' ] ?? '';
		$tokenInformation->givenName          = $claims[ 'given_name' ] ?? '';
		$tokenInformation->ip                 = $claims[ 'ipaddr' ] ?? '';
		$tokenInformation->scope              = $claims[ 'scp' ] ?? '';
		$tokenInformation->email              = $claims[ 'email' ] ?? '';
		$tokenInformation->preferred_username = $claims[ 'preferred_username' ] ?? '';
		$tokenInformation->uniqueName         = $claims[ 'unique_name' ] ?? '';
		$tokenInformation->upn                = $claims[ 'upn' ] ?? '';

		return $tokenInformation;
	}

}
