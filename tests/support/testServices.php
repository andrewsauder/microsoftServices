<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests\support;

/*
 * Each service class with its HTTP clients replaced by the mock clients of the running test. The file holds
 * several small classes, so graphTestCase loads it with require_once.
 */

final class testUser extends \andrewsauder\microsoftServices\user {

	use usesMockClients;
}

final class testCalendar extends \andrewsauder\microsoftServices\calendar {

	use usesMockClients;
}

final class testMail extends \andrewsauder\microsoftServices\mail {

	use usesMockClients;
}

final class testFiles extends \andrewsauder\microsoftServices\files {

	use usesMockClients;
}
