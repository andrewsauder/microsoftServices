<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests\support;

/** Replaces the HTTP clients of a service with the mock clients of the running test */
trait usesMockClients {

	protected function createHttpClient(): \GuzzleHttp\Client {
		return mockClients::$graph ?? throw new \LogicException( 'Set mockClients::$graph before you build the service' );
	}


	protected function createUploadHttpClient(): \GuzzleHttp\Client {
		return mockClients::$upload ?? throw new \LogicException( 'Set mockClients::$upload before you build the service' );
	}

}
