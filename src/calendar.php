<?php

namespace andrewsauder\microsoftServices;

use andrewsauder\microsoftServices\exceptions\serviceException;
use GuzzleHttp\Exception\GuzzleException;
use Microsoft\Graph\Exception\GraphException;

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
	 * @return \Microsoft\Graph\Model\Calendar[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getCalendars( string $calendarOwner ): array {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		try {

			$iterator = $graph->createCollectionRequest( 'GET', '/users/' . $emailAddress . '/calendars' )
			                  ->setReturnType( \Microsoft\Graph\Model\Calendar::class );
			$calendars  = $iterator->getPage();
			while( !$iterator->isEnd() ) {
				$calendars = array_merge( $calendars, $iterator->getPage() );
			}
			return $calendars;
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to get calendars: ' . $e->getMessage(), 500, $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to get calendars: ' . $e->getMessage(), $e->getCode(), $e );
		}
	}


	/**
	 * @param string $calendarOwner User id or username of a user whose calendars you want to fetch. Access token must have permission to access this account.
	 * @param string $calendarId Calendar id to get events from
	 *
	 * @return \Microsoft\Graph\Model\Event[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getCalendarEvents( string $calendarOwner, string $calendarId, \DateTimeImmutable $startDateTime, \DateTimeImmutable $endDateTime ): array {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		try {

			$iterator = $graph->createCollectionRequest( 'GET', '/users/' . $calendarOwner . '/calendars/'. $calendarId .'/calendarview?startdatetime='. $startDateTime->format('c').'&enddatetime='.  $endDateTime->format('c') .'&$orderby=start/dateTime' )
			                  ->setReturnType( \Microsoft\Graph\Model\Event::class );
			$events  = $iterator->getPage();
			while( !$iterator->isEnd() ) {
				$events = array_merge( $events, $iterator->getPage() );
			}
			return $events;
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to get calendar events: ' . $e->getMessage(), 500, $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to get calendar events: ' . $e->getMessage(), $e->getCode(), $e );
		}
	}


	/**
	 * @param string $calendarOwner User id or username of a user whose calendars you want to fetch. Access token must have permission to access this account.
	 * @param string $calendarId Calendar id to get events from
	 *
	 * @return \Microsoft\Graph\Model\Event[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getNCalendarEventsOccurringDuringOrAfter( string $calendarOwner, string $calendarId, \DateTimeImmutable $startDateTime, int $numberOfEventsToGet ): array {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		try {

			$iterator = $graph->createCollectionRequest( 'GET', '/users/' . $calendarOwner . '/calendars/'. $calendarId .'/events?$filter=start/dateTime ge \''.$startDateTime->format('c').'\' or end/dateTime ge \''.$startDateTime->format('c').'\'&$top='.$numberOfEventsToGet.'&$orderby=start/dateTime' )
			                  ->setReturnType( \Microsoft\Graph\Model\Event::class );
			$events  = $iterator->getPage();
			while( !$iterator->isEnd() ) {
				$events = array_merge( $events, $iterator->getPage() );
			}
			return $events;
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to get calendar events: ' . $e->getMessage(), 500, $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to get calendar events: ' . $e->getMessage(), $e->getCode(), $e );
		}
	}


}
