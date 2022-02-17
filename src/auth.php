<?php

namespace andrewsauder\microsoftServices;


use andrewsauder\microsoftServices\exceptions\serviceException;
use GuzzleHttp\Exception\GuzzleException;


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
	 * @return \League\OAuth2\Client\Token\AccessTokenInterface|\League\OAuth2\Client\Token\AccessToken
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getAccessToken( string $suppliedToken ) : \League\OAuth2\Client\Token\AccessTokenInterface|\League\OAuth2\Client\Token\AccessToken {
		$this->provider->defaultEndPointVersion = \TheNetworg\OAuth2\Client\Provider\Azure::ENDPOINT_VERSION_2_0;
		$this->provider->scope                  = $this->config->scope;

		try {
			/** @var \TheNetworg\OAuth2\Client\Token\AccessToken $token */
			return $this->provider->getAccessToken( 'jwt_bearer', [
				'scope'               => $this->provider->scope,
				'assertion'           => $suppliedToken,
				'requested_token_use' => 'on_behalf_of',
			] );
		}
		catch( \Exception $e ) {
			throw new \andrewsauder\microsoftServices\exceptions\serviceException( 'Microsoft authentication failed', 400, $e );
		}
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


	/**
	 * @return string
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getApplicationAccessToken() : string {
		//get application access token
		try {
			$guzzle = new \GuzzleHttp\Client();
			$url    = 'https://login.microsoftonline.com/' . $this->config->tenant . '/oauth2/token?api-version=1.0';
			$token  = json_decode( $guzzle->post( $url, [
				'form_params' => [
					'client_id'     => $this->config->clientId,
					'client_secret' => $this->config->clientSecret,
					'resource'      => 'https://graph.microsoft.com/',
					'grant_type'    => 'client_credentials',
				],
			] )->getBody()->getContents() );

			return $token->access_token;
		}
		catch( GuzzleException $e ) {
			throw new serviceException( $e->getMessage(), 500, $e );
		}
	}

}