<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests;

use andrewsauder\microsoftServices\calendar;
use andrewsauder\microsoftServices\tests\support\graphTestCase;
use andrewsauder\microsoftServices\tests\support\testCalendar;
use Microsoft\Graph\Generated\Models\Calendar as GraphCalendar;
use Microsoft\Graph\Generated\Models\Event;
use PHPUnit\Framework\Attributes\CoversClass;
use PHPUnit\Framework\Attributes\Test;

#[CoversClass( calendar::class )]
final class calendarTest extends graphTestCase {

	#[Test]
	public function getCalendarsReturnsEveryCalendar(): void {
		$this->queueJson( [ 'value' => [ [ 'id' => 'cal-1', 'name' => 'Calendar' ] ] ] );

		$calendars = ( new testCalendar( self::config(), self::TOKEN ) )->getCalendars( 'ann@contoso.example' );

		$this->assertSame( self::GRAPH . '/users/ann%40contoso.example/calendars', $this->path( 0 ) );
		$this->assertContainsOnlyInstancesOf( GraphCalendar::class, $calendars );
		$this->assertSame( 'Calendar', $calendars[ 0 ]->getName() );
	}


	#[Test]
	public function getCalendarEventsAsksForTheCalendarViewInOrder(): void {
		$this->queueJson( [
			'value'           => [ [ 'id' => 'event-1' ] ],
			'@odata.nextLink' => self::GRAPH . '/users/ann%40contoso.example/calendars/cal-1/calendarView?$skip=1'
		] );
		$this->queueJson( [ 'value' => [ [ 'id' => 'event-2' ] ] ] );

		$events = ( new testCalendar( self::config(), self::TOKEN ) )->getCalendarEvents(
			'ann@contoso.example',
			'cal-1',
			new \DateTimeImmutable( '2026-01-01T00:00:00+00:00' ),
			new \DateTimeImmutable( '2026-02-01T00:00:00+00:00' )
		);

		$this->assertRequestCount( 2 );
		$this->assertSame( self::GRAPH . '/users/ann%40contoso.example/calendars/cal-1/calendarView', $this->path( 0 ) );
		$query = $this->query( 0 );
		$this->assertSame( '2026-01-01T00:00:00+00:00', $query[ 'startDateTime' ] );
		$this->assertSame( '2026-02-01T00:00:00+00:00', $query[ 'endDateTime' ] );
		$this->assertSame( 'start/dateTime', $query[ '$orderby' ] );
		$this->assertSame( [ 'event-1', 'event-2' ], array_map( static fn( Event $event ) => $event->getId(), $events ) );
	}


	#[Test]
	public function getNCalendarEventsReturnsTheFirstPageOnly(): void {
		$this->queueJson( [
			'value'           => [ [ 'id' => 'event-1' ], [ 'id' => 'event-2' ] ],
			'@odata.nextLink' => self::GRAPH . '/users/ann%40contoso.example/calendars/cal-1/events?$skip=2'
		] );

		$events = ( new testCalendar( self::config(), self::TOKEN ) )->getNCalendarEventsOccurringDuringOrAfter(
			'ann@contoso.example',
			'cal-1',
			new \DateTimeImmutable( '2026-01-01T00:00:00+00:00' ),
			2
		);

		$this->assertRequestCount( 1 );
		$query = $this->query( 0 );
		$this->assertSame( '2', $query[ '$top' ] );
		$this->assertSame( "start/dateTime ge '2026-01-01T00:00:00+00:00' or end/dateTime ge '2026-01-01T00:00:00+00:00'", $query[ '$filter' ] );
		$this->assertCount( 2, $events );
	}

}
