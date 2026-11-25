<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests\support;

/**
 * The HTTP clients that the test services use in place of real ones. graphTestCase sets them before it builds a
 * service and clears them after each test.
 */
final class mockClients {

	public static ?\GuzzleHttp\Client $graph = null;

	public static ?\GuzzleHttp\Client $upload = null;

}
