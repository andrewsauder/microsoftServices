<?php

declare( strict_types=1 );

namespace andrewsauder\microsoftServices\tests;

use andrewsauder\microsoftServices\exceptions\serviceException;
use andrewsauder\microsoftServices\tests\support\graphTestCase;
use andrewsauder\microsoftServices\tests\support\testUser;
use andrewsauder\microsoftServices\user;
use Microsoft\Graph\Generated\Models\User as GraphUser;
use PHPUnit\Framework\Attributes\CoversClass;
use PHPUnit\Framework\Attributes\Test;

#[CoversClass( user::class )]
final class userTest extends graphTestCase {

	#[Test]
	public function allUsersInOrganizationFollowsEveryNextLink(): void {
		$this->queueJson( [
			'value'           => [ [ 'id' => '1', 'displayName' => 'Ann' ] ],
			'@odata.nextLink' => self::GRAPH . '/users?$skiptoken=page2'
		] );
		$this->queueJson( [ 'value' => [ [ 'id' => '2', 'displayName' => 'Bo' ] ] ] );

		$users = ( new testUser( self::config(), self::TOKEN ) )->allUsersInOrganization();

		$this->assertRequestCount( 2 );
		$this->assertSame( self::GRAPH . '/users', $this->path( 0 ) );
		$this->assertSame( 'Bearer ' . self::TOKEN, $this->request( 0 )->getHeaderLine( 'Authorization' ) );
		$this->assertSame( [ '$skiptoken' => 'page2' ], $this->query( 1 ) );
		$this->assertContainsOnlyInstancesOf( GraphUser::class, $users );
		$this->assertSame( [ 'Ann', 'Bo' ], array_map( static fn( GraphUser $user ) => $user->getDisplayName(), $users ) );
	}


	#[Test]
	public function getUserByUserPrincipalNameReturnsTheUser(): void {
		$this->queueJson( [ 'id' => '1', 'userPrincipalName' => 'ann@contoso.example' ] );

		$user = ( new testUser( self::config(), self::TOKEN ) )->getUserByUserPrincipalName( 'ann@contoso.example' );

		$this->assertSame( self::GRAPH . '/users/ann%40contoso.example', $this->path( 0 ) );
		$this->assertSame( 'ann@contoso.example', $user->getUserPrincipalName() );
	}


	#[Test]
	public function aMissingUserThrowsWithTheHttpStatus(): void {
		$this->queueError( 404, 'Request_ResourceNotFound', 'Resource does not exist.' );
		$this->expectErrorLog();

		try {
			( new testUser( self::config(), self::TOKEN ) )->getUserById( 'missing' );
			$this->fail( 'Expected a serviceException' );
		}
		catch( serviceException $e ) {
			$this->assertSame( 404, $e->getCode() );
			$this->assertSame( 'Failed to get user: Resource does not exist.', $e->getMessage() );
		}
	}


	#[Test]
	public function getUsersByFilterSendsTheFilter(): void {
		$this->queueJson( [ 'value' => [ [ 'id' => '1' ] ] ] );

		$users = ( new testUser( self::config(), self::TOKEN ) )->getUsersByFilter( "startswith(userPrincipalName,'ann')" );

		$this->assertSame( [ '$filter' => "startswith(userPrincipalName,'ann')" ], $this->query( 0 ) );
		$this->assertCount( 1, $users );
	}

}
