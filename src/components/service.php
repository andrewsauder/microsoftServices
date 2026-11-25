<?php

namespace andrewsauder\microsoftServices\components;

use andrewsauder\microsoftServices\exceptions\serviceException;
use GuzzleHttp\Exception\GuzzleException;
use Microsoft\Graph\Core\Authentication\GraphPhpLeagueAuthenticationProvider;
use Microsoft\Graph\Core\GraphClientFactory;
use Microsoft\Graph\Core\Tasks\PageIterator;
use Microsoft\Graph\Generated\Models\ODataErrors\ODataError;
use Microsoft\Graph\GraphRequestAdapter;
use Microsoft\Graph\GraphServiceClient;
use Microsoft\Kiota\Abstractions\ApiException;
use Microsoft\Kiota\Abstractions\Authentication\AuthenticationProvider;
use Microsoft\Kiota\Abstractions\Authentication\BaseBearerTokenAuthenticationProvider;
use Microsoft\Kiota\Abstractions\Serialization\Parsable;
use Microsoft\Kiota\Authentication\Oauth\ClientCredentialContext;
use Microsoft\Kiota\Authentication\Oauth\OnBehalfOfContext;

class service {

	protected \andrewsauder\microsoftServices\config $config;

	protected ?string $userAccessToken = null;

	protected GraphServiceClient $graph;


	/**
	 * @param string|null $userAccessToken Provide user token. If config.onBehalfOfFlow is enabled, the provided token will be exchanged for an access token for this API. If config.onBehalfOfFlow is not enabled, this token will be used
	 *                                     for the request. If no token is provided, an application token will be generated
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException 400 when the config is not valid
	 */
	public function __construct( \andrewsauder\microsoftServices\config $config, ?string $userAccessToken = null ) {
		$config->validate();

		$this->config          = $config;
		$this->userAccessToken = $userAccessToken;

		$this->graph = GraphServiceClient::createWithRequestAdapter( new GraphRequestAdapter( $this->createAuthenticationProvider(), $this->createHttpClient() ) );
	}


	/**
	 * Pick how the Graph client gets its access token:
	 * - no user token: an application token (client credentials)
	 * - a user token and config.onBehalfOfFlow: a token from the on-behalf-of exchange
	 * - a user token without config.onBehalfOfFlow: the user token as it is
	 */
	private function createAuthenticationProvider(): AuthenticationProvider {
		if( $this->userAccessToken===null ) {
			return new GraphPhpLeagueAuthenticationProvider( new ClientCredentialContext( $this->config->tenant, $this->config->clientId, $this->config->clientSecret ) );
		}

		if( $this->config->onBehalfOfFlow ) {
			return new GraphPhpLeagueAuthenticationProvider( new OnBehalfOfContext( $this->config->tenant, $this->config->clientId, $this->config->clientSecret, $this->userAccessToken ) );
		}

		return new BaseBearerTokenAuthenticationProvider( new staticAccessTokenProvider( $this->userAccessToken ) );
	}


	/**
	 * The HTTP client that sends Graph requests. Override it to send the requests somewhere else, for example
	 * to a mock handler in a test.
	 */
	protected function createHttpClient(): \GuzzleHttp\Client {
		return GraphClientFactory::create();
	}


	/**
	 * The HTTP client that PUTs file ranges to a pre-authenticated upload session URL. It sends no Graph
	 * headers and throws on an HTTP error status.
	 */
	protected function createUploadHttpClient(): \GuzzleHttp\Client {
		return new \GuzzleHttp\Client();
	}


	/**
	 * Wrap a failed Graph call in a serviceException. The code is the HTTP status of the response when there
	 * is one, else 500.
	 */
	protected function toServiceException( string $message, \Throwable $e ): serviceException {
		if( $e instanceof serviceException ) {
			return $e;
		}

		$code   = 500;
		$detail = $e->getMessage();

		if( $e instanceof ApiException ) {
			$code = $e->getResponseStatusCode() ?: 500;
			if( $e instanceof ODataError && $e->getPrimaryErrorMessage()!=='' ) {
				$detail = $e->getPrimaryErrorMessage();
			}
		}
		elseif( $e instanceof GuzzleException && (int)$e->getCode()>0 ) {
			$code = (int)$e->getCode();
		}

		error_log( (string)$e );

		return new serviceException( $detail==='' ? $message : $message . ': ' . $detail, $code, $e );
	}


	/**
	 * Every item of a paged collection response: the items of the first page, then the items of each page that
	 * the @odata.nextLink of the page before it points to.
	 *
	 * @param Parsable|null $firstPage A collection response, for example a UserCollectionResponse
	 *
	 * @return array<int, mixed>
	 * @throws \Exception
	 */
	protected function collectAll( ?Parsable $firstPage ): array {
		if( $firstPage===null ) {
			return [];
		}

		$items    = [];
		$iterator = new PageIterator( $firstPage, $this->graph->getRequestAdapter() );
		$iterator->iterate( static function( $item ) use ( &$items ): bool {
			$items[] = $item;
			return true;
		} );

		return $items;
	}


	/**
	 * PUT a file to an upload session URL, one range at a time, in order. The URL carries its own token, so the
	 * requests send no Authorization header.
	 *
	 * @return int The HTTP status of the last range: 200 or 201 when the upload is complete
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	protected function uploadFileInRanges( string $uploadUrl, string $filePath, int $size, string $name ): int {
		$handle = @fopen( $filePath, 'rb' );
		if( $handle===false ) {
			throw new serviceException( 'Unable to read ' . $name, 500 );
		}

		$client = $this->createUploadHttpClient();

		try {
			$status = 0;
			foreach( \andrewsauder\microsoftServices\mail\attachment::byteRanges( $size ) as [ $first, $last ] ) {
				$length = $last - $first + 1;
				$bytes  = stream_get_contents( $handle, $length, $first );
				if( $bytes===false || strlen( $bytes )!==$length ) {
					throw new serviceException( 'Unable to read ' . $name, 500 );
				}

				$response = $client->request( 'PUT', $uploadUrl, [
					'headers' => [
						'Content-Type'   => 'application/octet-stream',
						'Content-Length' => (string)$length,
						'Content-Range'  => \andrewsauder\microsoftServices\mail\attachment::contentRange( $first, $last, $size )
					],
					'body'    => $bytes
				] );
				$status   = $response->getStatusCode();
			}

			return $status;
		}
		catch( GuzzleException $e ) {
			throw $this->toServiceException( 'Failed to upload ' . $name, $e );
		}
		finally {
			fclose( $handle );
		}
	}

}
