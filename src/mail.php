<?php

namespace andrewsauder\microsoftServices;

use andrewsauder\microsoftServices\exceptions\serviceException;
use GuzzleHttp\Exception\GuzzleException;
use Microsoft\Graph\Exception\GraphException;

class mail extends \andrewsauder\microsoftServices\components\service {

	/** @var \andrewsauder\microsoftServices\mail\attachment[] */
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
	 * Queue a file to attach. The file is read when the message is created, so it must still exist then.
	 *
	 * createDraft() attaches each file in its own request, so a draft can carry files up to 150 MB each
	 * (the mailbox's message size limit still applies - 35 MB by default in Exchange Online). send() puts
	 * every file in its one sendMail request, which Graph rejects above about 4 MB in total; use
	 * createDraft() and sendDraft() for larger attachments.
	 *
	 * @param string $filePath
	 *
	 * @return string File name as it will be uploaded
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function addAttachment( string $filePath ): string {
		$attachment = \andrewsauder\microsoftServices\mail\attachment::fromFile( $filePath );

		$this->attachments[] = $attachment;

		return $attachment->name;
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
			$mailBody[ 'Message' ][ 'attachments' ] = array_map( static fn( \andrewsauder\microsoftServices\mail\attachment $attachment ) => $attachment->fileAttachmentBody(), $this->attachments );
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
	 * Create a draft message, then attach each queued file in its own request: a file up to 3 MB in one
	 * POST, a larger file (up to 150 MB) through an upload session. Because no single request carries
	 * every file, the total is limited only by the mailbox's message size limit, not by the Graph 4 MB
	 * request limit. If an attachment fails, the draft is deleted and the exception is thrown.
	 *
	 * The returned message is the draft as Graph created it, before the attachments were added. Its
	 * immutable id is unchanged by the attachments; pass it to sendDraft().
	 *
	 * @param string|string[] $to
	 * @param string          $subject
	 * @param string          $content
	 * @param string          $from
	 *
	 * @return \Microsoft\Graph\Model\Message
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function createDraft( string|array $to, string $subject, string $content, string $from, string|array $cc = [], string|array $bcc = [] ): \Microsoft\Graph\Model\Message {
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

		$ccRecipients = [];
		if( !is_array( $cc ) ) {
			$cc = [ $cc ];
		}
		foreach( $cc as $emailAddress ) {
			$ccRecipients[] = [
				'emailAddress' => [
					'address' => $emailAddress
				]
			];
		}

		$bccRecipients = [];
		if( !is_array( $bcc ) ) {
			$bcc = [ $bcc ];
		}
		foreach( $bcc as $emailAddress ) {
			$bccRecipients[] = [
				'emailAddress' => [
					'address' => $emailAddress
				]
			];
		}

		if( $from=='' ) {
			$from = $this->config->fromAddress;
		}

		$mailBody = [
			'subject'       => $subject,
			'body'          => [
				'contentType' => 'HTML',
				'content'     => $content
			],
			'from'          => [
				'emailAddress' => [
					'address' => $from
				]
			],
			'toRecipients'  => $toRecipients,
			'ccRecipients'  => $ccRecipients,
			'bccRecipients' => $bccRecipients
		];

		$headers = [
			'Prefer' => 'IdType="ImmutableId"'
		];

		try {
			$message = $graph->createRequest( 'POST', '/users/' . $from . '/messages' )
			                 ->setReturnType( \Microsoft\Graph\Model\Message::class )
			                 ->addHeaders( $headers )
			                 ->attachBody( $mailBody )
			                 ->execute();
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to create draft: ' . $e->getMessage(), 500, $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to create draft: ' . $e->getMessage(), $e->getCode(), $e );
		}

		if( count( $this->attachments )>0 ) {
			$messageId = (string)$message->getId();
			try {
				foreach( $this->attachments as $attachment ) {
					$this->attachToMessage( $graph, $from, $messageId, $attachment );
				}
			}
			catch( serviceException $e ) {
				$this->deleteMessageQuietly( $graph, $from, $messageId );
				throw $e;
			}
		}

		return $message;
	}


	/**
	 * Attach one file to an existing draft message: one POST for a small file, an upload session for a
	 * large one
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	private function attachToMessage( \Microsoft\Graph\Graph $graph, string $from, string $messageId, \andrewsauder\microsoftServices\mail\attachment $attachment ): void {
		$attachmentsUrl = '/users/' . $from . '/messages/' . $messageId . '/attachments';

		if( !$attachment->usesUploadSession() ) {
			try {
				$graph->createRequest( 'POST', $attachmentsUrl )
				      ->attachBody( $attachment->fileAttachmentBody() )
				      ->execute();
			}
			catch( GraphException $e ) {
				error_log( $e );
				throw new serviceException( 'Failed to attach ' . $attachment->name . ': ' . $e->getMessage(), 500, $e );
			}
			catch( GuzzleException $e ) {
				error_log( $e );
				throw new serviceException( 'Failed to attach ' . $attachment->name . ': ' . $e->getMessage(), $e->getCode(), $e );
			}
			return;
		}

		try {
			/** @var \Microsoft\Graph\Model\UploadSession $uploadSession */
			$uploadSession = $graph->createRequest( 'POST', $attachmentsUrl . '/createUploadSession' )
			                       ->setReturnType( \Microsoft\Graph\Model\UploadSession::class )
			                       ->attachBody( $attachment->uploadSessionBody() )
			                       ->execute();
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to start the upload of ' . $attachment->name . ': ' . $e->getMessage(), 500, $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to start the upload of ' . $attachment->name . ': ' . $e->getMessage(), $e->getCode(), $e );
		}

		$uploadUrl = (string)$uploadSession->getUploadUrl();
		if( $uploadUrl==='' ) {
			throw new serviceException( 'Graph did not return an upload URL for ' . $attachment->name, 500 );
		}

		$this->uploadInRanges( $uploadUrl, $attachment );
	}


	/**
	 * PUT the file to an upload session URL, one range at a time, in order. The URL carries its own
	 * token, so the requests send no Authorization header. Graph answers 200 to each range and
	 * 201 Created to the last one, which completes the attachment.
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	private function uploadInRanges( string $uploadUrl, \andrewsauder\microsoftServices\mail\attachment $attachment ): void {
		$handle = @fopen( $attachment->path, 'rb' );
		if( $handle===false ) {
			throw new serviceException( 'Unable to read the attachment ' . $attachment->name, 500 );
		}

		$client = new \GuzzleHttp\Client();

		try {
			$status = 0;
			foreach( \andrewsauder\microsoftServices\mail\attachment::byteRanges( $attachment->size ) as [ $first, $last ] ) {
				$length = $last - $first + 1;
				$bytes  = stream_get_contents( $handle, $length, $first );
				if( $bytes===false || strlen( $bytes )!==$length ) {
					throw new serviceException( 'Unable to read the attachment ' . $attachment->name, 500 );
				}

				$response = $client->request( 'PUT', $uploadUrl, [
					'headers' => [
						'Content-Type'   => 'application/octet-stream',
						'Content-Length' => (string)$length,
						'Content-Range'  => \andrewsauder\microsoftServices\mail\attachment::contentRange( $first, $last, $attachment->size )
					],
					'body'    => $bytes
				] );
				$status   = $response->getStatusCode();
			}

			if( $status!==201 ) {
				throw new serviceException( 'Graph did not complete the upload of ' . $attachment->name . ' (HTTP ' . $status . ')', 500 );
			}
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to upload ' . $attachment->name . ': ' . $e->getMessage(), $e->getCode() ?: 500, $e );
		}
		finally {
			fclose( $handle );
		}
	}


	/** Remove a draft that a failed attachment left half built. A failure here is logged, not thrown. */
	private function deleteMessageQuietly( \Microsoft\Graph\Graph $graph, string $from, string $messageId ): void {
		try {
			$graph->createRequest( 'DELETE', '/users/' . $from . '/messages/' . $messageId )->execute();
		}
		catch( \Throwable $e ) {
			error_log( 'Failed to delete the draft ' . $messageId . ' after an attachment failed: ' . $e->getMessage() );
		}
	}


	/**
	 * @param string $messageId
	 * @param string $from
	 *
	 * @return \Microsoft\Graph\Model\Message
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getMessage( string $messageId, string $from ): \Microsoft\Graph\Model\Message {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//get file list
		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		$headers = [
			'Prefer' => 'IdType="ImmutableId"'
		];

		try {
			$message = $graph->createRequest( 'GET', '/users/' . $from . '/messages/' . $messageId )
			                 ->setReturnType( \Microsoft\Graph\Model\Message::class )
			                 ->addHeaders( $headers )
			                 ->execute();
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to create draft: ' . $e->getMessage(), 500, $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to create draft: ' . $e->getMessage(), $e->getCode(), $e );
		}

		return $message;
	}


	public function sendDraft( string $messageId, string $from ): \Microsoft\Graph\Http\GraphResponse {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//get file list
		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		try {
			$response = $graph->createRequest( 'POST', '/users/' . $from . '/messages/' . $messageId . '/send' )
			                  ->execute();
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to create draft: ' . $e->getMessage(), 500, $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to create draft: ' . $e->getMessage(), $e->getCode(), $e );
		}

		return $response;
	}


	/**
	 * Create a draft reply to an existing message. The draft is created in the SAME conversation as the
	 * original message and inherits its RFC threading headers (In-Reply-To / References) and a "RE:"
	 * subject automatically - so those must NOT be set by hand. Recipients and body are set separately
	 * (see replyIntoThread()).
	 *
	 * @param string $messageId Immutable id of the message being replied to (the thread anchor)
	 * @param string $from      Mailbox that owns the anchor message and will send the reply
	 *
	 * @return \Microsoft\Graph\Model\Message The created draft, in the same conversation as $messageId
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function createReplyDraft( string $messageId, string $from ): \Microsoft\Graph\Model\Message {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		if( $from=='' ) {
			$from = $this->config->fromAddress;
		}

		//request immutable ids so the returned draft id survives the draft->Sent Items move and is safe
		//to store as the next thread anchor
		$headers = [
			'Prefer' => 'IdType="ImmutableId"'
		];

		try {
			$message = $graph->createRequest( 'POST', '/users/' . $from . '/messages/' . $messageId . '/createReply' )
			                 ->setReturnType( \Microsoft\Graph\Model\Message::class )
			                 ->addHeaders( $headers )
			                 ->execute();
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to create reply draft: ' . $e->getMessage(), 500, $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to create reply draft: ' . $e->getMessage(), $e->getCode(), $e );
		}

		return $message;
	}


	/**
	 * Reply into an existing conversation as a single grouped message: create a reply draft anchored on
	 * $anchorMessageId, set its recipients and body, then send it. Because createReply inherits the
	 * conversation and threading headers, the message lands in the same thread as the anchor for every
	 * recipient.
	 *
	 * The subject is intentionally not settable here - the reply inherits the anchor's subject and Graph
	 * prefixes "RE:" automatically.
	 *
	 * @param string          $from            Mailbox that owns the anchor and sends the reply
	 * @param string          $anchorMessageId Immutable id of the last message sent in the thread
	 * @param string|string[] $to              To recipients
	 * @param string|string[] $bcc             Bcc recipients
	 * @param string          $htmlBody        HTML message body
	 * @param string|string[] $cc              Cc recipients
	 *
	 * @return \Microsoft\Graph\Model\Message The sent message; read ->getId() (the next anchor) and ->getConversationId()
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function replyIntoThread( string $from, string $anchorMessageId, string|array $to, string|array $bcc, string $htmlBody, string|array $cc = [] ): \Microsoft\Graph\Model\Message {
		if( $from=='' ) {
			$from = $this->config->fromAddress;
		}

		//1) create the reply draft in the same conversation as the anchor
		$draft   = $this->createReplyDraft( $anchorMessageId, $from );
		$draftId = $draft->getId();

		//get application access token for the remaining calls
		$accessToken = $this->getMicrosoftAccessToken();

		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		//2) set recipients + body on the draft (subject and threading headers are inherited from the anchor)
		$patchBody = [
			'toRecipients'  => $this->buildRecipients( $to ),
			'ccRecipients'  => $this->buildRecipients( $cc ),
			'bccRecipients' => $this->buildRecipients( $bcc ),
			'body'          => [
				'contentType' => 'HTML',
				'content'     => $htmlBody
			]
		];

		$headers = [
			'Prefer' => 'IdType="ImmutableId"'
		];

		try {
			$message = $graph->createRequest( 'PATCH', '/users/' . $from . '/messages/' . $draftId )
			                 ->setReturnType( \Microsoft\Graph\Model\Message::class )
			                 ->addHeaders( $headers )
			                 ->attachBody( $patchBody )
			                 ->execute();
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to update reply draft: ' . $e->getMessage(), 500, $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to update reply draft: ' . $e->getMessage(), $e->getCode(), $e );
		}

		//3) send the draft
		try {
			$graph->createRequest( 'POST', '/users/' . $from . '/messages/' . $draftId . '/send' )
			      ->execute();
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to send reply: ' . $e->getMessage(), 500, $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to send reply: ' . $e->getMessage(), $e->getCode(), $e );
		}

		//the patched draft carries the immutable id (still valid after the move to Sent Items) and the
		//conversationId inherited from the anchor
		return $message;
	}


	/**
	 * Fallback anchor discovery: return the most recent message in a conversation, or null when none is
	 * found. Used when a stored last_message_id is missing but the conversationId is known.
	 *
	 * @param string $from           Mailbox to search
	 * @param string $conversationId Graph conversationId of the thread
	 *
	 * @return \Microsoft\Graph\Model\Message|null
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getLatestMessageInConversation( string $from, string $conversationId ): ?\Microsoft\Graph\Model\Message {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		if( $from=='' ) {
			$from = $this->config->fromAddress;
		}

		//escape single quotes for the OData string literal, then url-encode the query values
		$escapedConversationId = str_replace( "'", "''", $conversationId );
		$filter                = rawurlencode( "conversationId eq '" . $escapedConversationId . "'" );
		$orderBy               = rawurlencode( 'receivedDateTime desc' );
		$url                   = '/users/' . $from . '/messages?$filter=' . $filter . '&$select=id,conversationId&$orderby=' . $orderBy . '&$top=1';

		$headers = [
			'Prefer' => 'IdType="ImmutableId"'
		];

		try {
			$messages = $graph->createRequest( 'GET', $url )
			                  ->setReturnType( \Microsoft\Graph\Model\Message::class )
			                  ->addHeaders( $headers )
			                  ->execute();
		}
		catch( GraphException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to look up conversation: ' . $e->getMessage(), 500, $e );
		}
		catch( GuzzleException $e ) {
			error_log( $e );
			throw new serviceException( 'Failed to look up conversation: ' . $e->getMessage(), $e->getCode(), $e );
		}

		if( is_array( $messages ) && count( $messages )>0 ) {
			return $messages[ 0 ];
		}

		return null;
	}


	/**
	 * Build a Graph recipient array from a single address or a list of addresses. Empty values are skipped.
	 *
	 * @param string|string[] $addresses
	 *
	 * @return array
	 */
	private function buildRecipients( string|array $addresses ): array {
		if( !is_array( $addresses ) ) {
			$addresses = [ $addresses ];
		}

		$recipients = [];
		foreach( $addresses as $emailAddress ) {
			if( $emailAddress===null || $emailAddress==='' ) {
				continue;
			}
			$recipients[] = [
				'emailAddress' => [
					'address' => $emailAddress
				]
			];
		}

		return $recipients;
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
		return $this->listMessages( '/users/' . $emailAddress . '/mailFolders/' . $mailFolderId . '/messages', $preferHeader );
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
			                         ->addHeaders( [ 'Prefer' => 'outlook.body-content-type="' . $preferHeader->value . '", IdType="ImmutableId"' ] );
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
			$attachmentIterator = $graph->createCollectionRequest( "GET", '/users/' . $emailAddress . '/messages/' . $messageId . "/attachments" )
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
	 * @return \Microsoft\Graph\Http\GraphResponse
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function deleteMessage( string $emailAddress, string $messageId ): \Microsoft\Graph\Http\GraphResponse {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//get file list
		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		try {
			$response = $graph->createRequest( 'DELETE', '/users/' . $emailAddress . '/messages/' . $messageId )->execute();
			return $response;
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


	/**
	 * @param string $emailAddress
	 * @param string $messageId
	 * @param string $mailFolderId
	 *
	 * @return \Microsoft\Graph\Http\GraphResponse
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function moveMessage( string $emailAddress, string $messageId, string $mailFolderId ): \Microsoft\Graph\Http\GraphResponse {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//get file list
		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		try {
			$body = [
				'destinationId' => $mailFolderId,
			];
			$response = $graph->createRequest( 'POST', '/users/' . $emailAddress . '/messages/' . $messageId .'/move' )->attachBody( $body )->execute();
			return $response;
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
