<?php

namespace andrewsauder\microsoftServices;

use Microsoft\Graph\Generated\Users\Item\Calendars\Item\CalendarView\CalendarViewRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\Generated\Users\Item\Calendars\Item\Events\EventsRequestBuilderGetRequestConfiguration;

class calendar extends \andrewsauder\microsoftServices\components\service {


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
	 * @param string $calendarOwner User id or username of a user whose calendars you want to fetch. Access token must have permission to access this account.
	 *
	 * @return \Microsoft\Graph\Generated\Models\Calendar[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getCalendars( string $calendarOwner ): array {
		try {
			return $this->collectAll( $this->graph->users()->byUserId( $calendarOwner )->calendars()->get()->wait() );
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to get calendars', $e );
		}
	}


	/**
	 * Every event that occurs between the two dates, in order of start time. Recurring events are expanded into their
	 * occurrences.
	 *
	 * @param string $calendarOwner User id or username of a user whose calendars you want to fetch. Access token must have permission to access this account.
	 * @param string $calendarId    Calendar id to get events from
	 *
	 * @return \Microsoft\Graph\Generated\Models\Event[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getCalendarEvents( string $calendarOwner, string $calendarId, \DateTimeImmutable $startDateTime, \DateTimeImmutable $endDateTime ): array {
		$requestConfiguration = new CalendarViewRequestBuilderGetRequestConfiguration(
			queryParameters: CalendarViewRequestBuilderGetRequestConfiguration::createQueryParameters(
				endDateTime  : $endDateTime->format( 'c' ),
				orderby      : [ 'start/dateTime' ],
				startDateTime: $startDateTime->format( 'c' )
			)
		);

		try {
			return $this->collectAll(
				$this->graph->users()->byUserId( $calendarOwner )->calendars()->byCalendarId( $calendarId )->calendarView()->get( $requestConfiguration )->wait()
			);
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to get calendar events', $e );
		}
	}


	/**
	 * At most $numberOfEventsToGet events that start or end at or after $startDateTime, in order of start time
	 *
	 * @param string $calendarOwner User id or username of a user whose calendars you want to fetch. Access token must have permission to access this account.
	 * @param string $calendarId    Calendar id to get events from
	 *
	 * @return \Microsoft\Graph\Generated\Models\Event[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getNCalendarEventsOccurringDuringOrAfter( string $calendarOwner, string $calendarId, \DateTimeImmutable $startDateTime, int $numberOfEventsToGet ): array {
		$start = $startDateTime->format( 'c' );

		$requestConfiguration = new EventsRequestBuilderGetRequestConfiguration(
			queryParameters: EventsRequestBuilderGetRequestConfiguration::createQueryParameters(
				filter : "start/dateTime ge '" . $start . "' or end/dateTime ge '" . $start . "'",
				orderby: [ 'start/dateTime' ],
				top    : $numberOfEventsToGet
			)
		);

		try {
			$page = $this->graph->users()->byUserId( $calendarOwner )->calendars()->byCalendarId( $calendarId )->events()->get( $requestConfiguration )->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to get calendar events', $e );
		}

		//the first page holds at most $top events; the next pages are past the number asked for
		return $page?->getValue() ?? [];
	}

}
