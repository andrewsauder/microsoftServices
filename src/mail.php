<?php

namespace andrewsauder\microsoftServices;

use andrewsauder\microsoftServices\exceptions\serviceException;
use GuzzleHttp\Exception\GuzzleException;
use Microsoft\Graph\Exception\GraphException;

class mail extends \andrewsauder\microsoftServices\components\service {

	private array $attachments = [];


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
	 * @param string $filePath
	 *
	 * @return string File name as it will be uploaded
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function addAttachment( string $filePath ): string {
		if( !file_exists( $filePath ) ) {
			throw new serviceException( 'File does not exist', 500 );
		}

		$fileName     = basename( $filePath );
		$fileContents = file_get_contents( $filePath );
		$mimeType     = mime_content_type( $filePath );

		$this->attachments[] = [
			'@odata.type'  => '#microsoft.graph.fileAttachment',
			'name'         => $fileName,
			'contentType'  => $mimeType,
			'contentBytes' => base64_encode( $fileContents )
		];

		return $fileName;
	}


	/**
	 * @param string|string[] $to
	 * @param string          $subject
	 * @param string          $content
	 * @param string          $from
	 *
	 * @return \Microsoft\Graph\Http\GraphResponse
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function send( string|array $to, string $subject, string $content, string $from = '' ): \Microsoft\Graph\Http\GraphResponse {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//get file list
		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		$toRecipients = [];
		if( !is_array( $to ) ) {
			$to = [ $to ];
		}
		foreach( $to as $emailAddress ) {
			$toRecipients[] = [
				'emailAddress' => [
					'address' => $emailAddress
				]
			];
		}

		if( $from=='' ) {
			$from = $this->config->fromAddress;
		}

		$mailBody = [
			'Message' => [
				'subject'      => $subject,
				'body'         => [
					'contentType' => 'HTML',
					'content'     => $content
				],
				'from'         => [
					'emailAddress' => [
						'address' => $from
					]
				],
				'toRecipients' => $toRecipients
			]
		];

		if( count( $this->attachments )>0 ) {
			$mailBody[ 'Message' ][ 'attachments' ] = $this->attachments;
		}

		try {
			$response = $graph->createRequest( 'POST', '/users/' . $from . '/sendMail' )->attachBody( $mailBody )->execute();
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to send email: ' . $e->getMessage(), 500, $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to send email: ' . $e->getMessage(), $e->getCode(), $e );
		}
		return $response;
	}


	/**
	 * @param string                                              $emailAddress
	 * @param \andrewsauder\microsoftServices\mail\headers\prefer $preferHeader [optional] Defaults to html
	 *
	 * @return \Microsoft\Graph\Model\Message[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getAllMessages( string $emailAddress, \andrewsauder\microsoftServices\mail\headers\prefer $preferHeader = \andrewsauder\microsoftServices\mail\headers\prefer::HTML ): array {
		return $this->listMessages( '/users/' . $emailAddress . '/messages', $preferHeader );
	}


	/**
	 * @param string                                              $emailAddress
	 * @param string                                              $mailFolderId
	 * @param \andrewsauder\microsoftServices\mail\headers\prefer $preferHeader [optional] Defaults to html
	 *
	 * @return \Microsoft\Graph\Model\Message[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getMessagesInFolder( string $emailAddress, string $mailFolderId, \andrewsauder\microsoftServices\mail\headers\prefer $preferHeader = \andrewsauder\microsoftServices\mail\headers\prefer::HTML ): array {
		return $this->listMessages( '/users/' . $emailAddress . '/mailFolder/' . $mailFolderId . '/messages', $preferHeader );
	}


	/**
	 * @param string $emailAddress
	 * @param bool   $includeHiddenFolders
	 *
	 * @return \Microsoft\Graph\Model\MailFolder[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getFolders( string $emailAddress, bool $includeHiddenFolders = false ): array {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//get file list
		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		$url = '/users/' . $emailAddress . '/mailFolders';
		if( $includeHiddenFolders ) {
			$url .= '?includeHiddenFolders=true';
		}

		try {
			$iterator = $graph->createCollectionRequest( 'GET', $url )
				->setReturnType( \Microsoft\Graph\Model\MailFolder::class );
			$folders  = $iterator->getPage();
			while( !$iterator->isEnd() ) {
				$folders = array_merge( $folders, $iterator->getPage() );
			}
			return $folders;
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to send email: ' . $e->getMessage(), 500, $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to send email: ' . $e->getMessage(), $e->getCode(), $e );
		}
	}


	/**
	 * @param string                                              $url
	 * @param \andrewsauder\microsoftServices\mail\headers\prefer $preferHeader [optional] Defaults to html
	 *
	 * @return \Microsoft\Graph\Model\Message[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	private function listMessages( string $url, \andrewsauder\microsoftServices\mail\headers\prefer $preferHeader = \andrewsauder\microsoftServices\mail\headers\prefer::HTML ): array {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//get file list
		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		try {
			$messageIterator = $graph->createCollectionRequest( 'GET', $url )
				->setReturnType( \Microsoft\Graph\Model\Message::class )
				->addHeaders( [ 'Prefer' => "outlook.body-content-type='" . $preferHeader->value . "'" ] );
			$messages        = $messageIterator->getPage();
			while( !$messageIterator->isEnd() ) {
				$messages = array_merge( $messages, $messageIterator->getPage() );
			}
			return $messages;
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to send email: ' . $e->getMessage(), 500, $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to send email: ' . $e->getMessage(), $e->getCode(), $e );
		}
	}


	/**
	 * @param string $emailAddress
	 * @param string $messageId
	 *
	 * @return \Microsoft\Graph\Model\Attachment[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getAttachments( string $emailAddress, string $messageId ): array {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//get file list
		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		try {
			$attachmentIterator = $graph->createCollectionRequest( "GET", '/users/'.$emailAddress.'/messages/' . $messageId . "/attachments" )
				->setReturnType( \Microsoft\Graph\Model\Attachment::class );

			$attachments = $attachmentIterator->getPage();

			while( !$attachmentIterator->isEnd() ) {
				$attachments = array_merge( $attachments, $attachmentIterator->getPage() );
			}

			return $attachments;
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to send email: ' . $e->getMessage(), 500, $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to send email: ' . $e->getMessage(), $e->getCode(), $e );
		}
	}


	/**
	 * @param string $emailAddress
	 * @param string $messageId
	 *
	 * @return void
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function deleteMessage( string $emailAddress, string $messageId ): void {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//get file list
		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		try {
			$response = $graph->createCollectionRequest( 'DELETE', '/users/' . $emailAddress . '/messages/' . $messageId );
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to delete email message: ' . $e->getMessage(), 500, $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to delete email message: ' . $e->getMessage(), $e->getCode(), $e );
		}
	}

}