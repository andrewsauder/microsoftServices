<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests\support;

use andrewsauder\microsoftServices\config;
use GuzzleHttp\Client;
use GuzzleHttp\Handler\MockHandler;
use GuzzleHttp\HandlerStack;
use GuzzleHttp\Middleware;
use GuzzleHttp\Psr7\Response;
use PHPUnit\Framework\TestCase;
use Psr\Http\Message\RequestInterface;

require_once __DIR__ . '/testServices.php';

/**
 * Builds services that send their requests to a Guzzle mock handler instead of Microsoft Graph. The config turns
 * the on-behalf-of flow off and the service gets a user token, so the services request no token either.
 *
 * Queue each response in the order the service sends its requests, call the service, then read the requests back
 * from the history.
 */
abstract class graphTestCase extends TestCase {

	protected const TOKEN = 'test-access-token';

	protected const GRAPH = 'https://graph.microsoft.com/v1.0';

	protected MockHandler $mockHandler;

	/** @var array<int, array{request: RequestInterface}> */
	protected array $history = [];


	protected function setUp(): void {
		$this->mockHandler = new MockHandler();
		$this->history     = [];

		$stack = HandlerStack::create( $this->mockHandler );
		$stack->push( Middleware::history( $this->history ) );

		//the Graph client reads each status itself; the upload client throws on an error status, like a default client
		mockClients::$graph  = new Client( [ 'handler' => $stack, 'http_errors' => false ] );
		mockClients::$upload = new Client( [ 'handler' => $stack ] );
	}


	protected function tearDown(): void {
		mockClients::$graph  = null;
		mockClients::$upload = null;
	}


	protected static function config(): config {
		$config                 = new config();
		$config->clientId       = 'client-id';
		$config->clientSecret   = 'client-secret';
		$config->tenant         = 'contoso.example';
		$config->driveId        = 'drive-1';
		$config->fromAddress    = 'noreply@contoso.example';
		$config->onBehalfOfFlow = false;

		return $config;
	}


	/** Queue a JSON response */
	protected function queueJson( array $body, int $status = 200 ): void {
		$this->mockHandler->append( new Response( $status, [ 'Content-Type' => 'application/json' ], json_encode( $body, JSON_THROW_ON_ERROR ) ) );
	}


	/** Queue a response with no body */
	protected function queueEmpty( int $status = 204 ): void {
		$this->mockHandler->append( new Response( $status ) );
	}


	/** Queue a Graph error response */
	protected function queueError( int $status, string $code, string $message ): void {
		$this->queueJson( [ 'error' => [ 'code' => $code, 'message' => $message ] ], $status );
	}


	protected function request( int $index ): RequestInterface {
		$this->assertArrayHasKey( $index, $this->history, 'The service sent fewer requests than expected' );

		return $this->history[ $index ][ 'request' ];
	}


	protected function assertRequestCount( int $count ): void {
		$this->assertCount( $count, $this->history );
		$this->assertSame( 0, $this->mockHandler->count(), 'Some queued responses were not used' );
	}


	/** The URL of a request without its query */
	protected function path( int $index ): string {
		$uri = $this->request( $index )->getUri();

		return $uri->getScheme() . '://' . $uri->getHost() . $uri->getPath();
	}


	/**
	 * The decoded query parameters of a request. Unlike parse_str(), this keeps '$' and '.' in the names.
	 *
	 * @return array<string, string>
	 */
	protected function query( int $index ): array {
		$query = $this->request( $index )->getUri()->getQuery();
		if( $query==='' ) {
			return [];
		}

		$parameters = [];
		foreach( explode( '&', $query ) as $pair ) {
			[ $name, $value ] = array_pad( explode( '=', $pair, 2 ), 2, '' );
			$parameters[ rawurldecode( $name ) ] = rawurldecode( $value );
		}

		return $parameters;
	}


	/** The decoded JSON body of a request */
	protected function jsonBody( int $index ): array {
		$body = (string)$this->request( $index )->getBody();

		return json_decode( $body, true, 512, JSON_THROW_ON_ERROR );
	}

}
