<?php

namespace andrewsauder\microsoftServices;

use andrewsauder\microsoftServices\exceptions\serviceException;
use andrewsauder\microsoftServices\mail\attachment;
use andrewsauder\microsoftServices\mail\headers\prefer;
use Microsoft\Graph\Generated\Models\BodyType;
use Microsoft\Graph\Generated\Models\EmailAddress;
use Microsoft\Graph\Generated\Models\ItemBody;
use Microsoft\Graph\Generated\Models\Message;
use Microsoft\Graph\Generated\Models\Recipient;
use Microsoft\Graph\Generated\Users\Item\MailFolders\Item\Messages\MessagesRequestBuilderGetRequestConfiguration as FolderMessagesRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\Generated\Users\Item\MailFolders\MailFoldersRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\Generated\Users\Item\Messages\Item\CreateReply\CreateReplyPostRequestBody;
use Microsoft\Graph\Generated\Users\Item\Messages\Item\CreateReply\CreateReplyRequestBuilderPostRequestConfiguration;
use Microsoft\Graph\Generated\Users\Item\Messages\Item\MessageItemRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\Generated\Users\Item\Messages\Item\MessageItemRequestBuilderPatchRequestConfiguration;
use Microsoft\Graph\Generated\Users\Item\Messages\Item\Move\MovePostRequestBody;
use Microsoft\Graph\Generated\Users\Item\Messages\MessagesRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\Generated\Users\Item\Messages\MessagesRequestBuilderPostRequestConfiguration;
use Microsoft\Graph\Generated\Users\Item\SendMail\SendMailPostRequestBody;

class mail extends \andrewsauder\microsoftServices\components\service {

	/** The Prefer header that asks Graph for immutable message ids */
	private const IMMUTABLE_ID = 'IdType="ImmutableId"';

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
		$attachment = attachment::fromFile( $filePath );

		$this->attachments[] = $attachment;

		return $attachment->name;
	}


	/**
	 * Send a message in one sendMail request. Graph saves it to the Sent Items folder of $from.
	 *
	 * @param string|string[] $to
	 * @param string          $subject
	 * @param string          $content HTML message body
	 * @param string          $from    Defaults to config.fromAddress
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException when Graph does not accept the message
	 */
	public function send( string|array $to, string $subject, string $content, string $from = '' ): void {
		if( $from=='' ) {
			$from = $this->config->fromAddress;
		}

		$message = new Message();
		$message->setSubject( $subject );
		$message->setBody( $this->htmlBody( $content ) );
		$message->setFrom( $this->recipient( $from ) );
		$message->setToRecipients( $this->buildRecipients( $to ) );

		if( count( $this->attachments )>0 ) {
			$message->setAttachments( array_map( static fn( attachment $attachment ) => $attachment->toFileAttachment(), $this->attachments ) );
		}

		$requestBody = new SendMailPostRequestBody();
		$requestBody->setMessage( $message );

		try {
			$this->graph->users()->byUserId( $from )->sendMail()->post( $requestBody )->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to send email', $e );
		}
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
	 * @param string          $content HTML message body
	 * @param string          $from    Defaults to config.fromAddress
	 * @param string|string[] $cc
	 * @param string|string[] $bcc
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function createDraft( string|array $to, string $subject, string $content, string $from, string|array $cc = [], string|array $bcc = [] ): Message {
		if( $from=='' ) {
			$from = $this->config->fromAddress;
		}

		$draft = new Message();
		$draft->setSubject( $subject );
		$draft->setBody( $this->htmlBody( $content ) );
		$draft->setFrom( $this->recipient( $from ) );
		$draft->setToRecipients( $this->buildRecipients( $to ) );
		$draft->setCcRecipients( $this->buildRecipients( $cc ) );
		$draft->setBccRecipients( $this->buildRecipients( $bcc ) );

		try {
			$message = $this->graph->users()->byUserId( $from )->messages()->post( $draft, new MessagesRequestBuilderPostRequestConfiguration( headers: [ 'Prefer' => self::IMMUTABLE_ID ] ) )->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to create draft', $e );
		}

		if( !$message instanceof Message ) {
			throw new serviceException( 'Graph did not return the draft it created', 500 );
		}

		if( count( $this->attachments )>0 ) {
			$messageId = (string)$message->getId();
			try {
				foreach( $this->attachments as $attachment ) {
					$this->attachToMessage( $from, $messageId, $attachment );
				}
			}
			catch( serviceException $e ) {
				$this->deleteMessageQuietly( $from, $messageId );
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
	private function attachToMessage( string $from, string $messageId, attachment $attachment ): void {
		$attachments = $this->graph->users()->byUserId( $from )->messages()->byMessageId( $messageId )->attachments();

		if( !$attachment->usesUploadSession() ) {
			try {
				$attachments->post( $attachment->toFileAttachment() )->wait();
			}
			catch( \Throwable $e ) {
				throw $this->toServiceException( 'Failed to attach ' . $attachment->name, $e );
			}
			return;
		}

		try {
			$uploadSession = $attachments->createUploadSession()->post( $attachment->toUploadSessionRequestBody() )->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to start the upload of ' . $attachment->name, $e );
		}

		$uploadUrl = (string)$uploadSession?->getUploadUrl();
		if( $uploadUrl==='' ) {
			throw new serviceException( 'Graph did not return an upload URL for ' . $attachment->name, 500 );
		}

		//Graph answers 200 to each range and 201 Created to the last one, which completes the attachment
		$status = $this->uploadFileInRanges( $uploadUrl, $attachment->path, $attachment->size, $attachment->name );
		if( $status!==201 ) {
			throw new serviceException( 'Graph did not complete the upload of ' . $attachment->name . ' (HTTP ' . $status . ')', 500 );
		}
	}


	/** Remove a draft that a failed attachment left half built. A failure here is logged, not thrown. */
	private function deleteMessageQuietly( string $from, string $messageId ): void {
		try {
			$this->graph->users()->byUserId( $from )->messages()->byMessageId( $messageId )->delete()->wait();
		}
		catch( \Throwable $e ) {
			error_log( 'Failed to delete the draft ' . $messageId . ' after an attachment failed: ' . $e->getMessage() );
		}
	}


	/**
	 * @param string $messageId
	 * @param string $from
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getMessage( string $messageId, string $from ): Message {
		try {
			$message = $this->graph->users()->byUserId( $from )->messages()->byMessageId( $messageId )
			                       ->get( new MessageItemRequestBuilderGetRequestConfiguration( headers: [ 'Prefer' => self::IMMUTABLE_ID ] ) )
			                       ->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to get message', $e );
		}

		if( !$message instanceof Message ) {
			throw new serviceException( 'Failed to get message', 404 );
		}

		return $message;
	}


	/**
	 * Send a draft message. Graph moves it to the Sent Items folder of $from.
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException when Graph does not send the draft
	 */
	public function sendDraft( string $messageId, string $from ): void {
		try {
			$this->graph->users()->byUserId( $from )->messages()->byMessageId( $messageId )->send()->post()->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to send draft', $e );
		}
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
	 * @return \Microsoft\Graph\Generated\Models\Message The created draft, in the same conversation as $messageId
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function createReplyDraft( string $messageId, string $from ): Message {
		if( $from=='' ) {
			$from = $this->config->fromAddress;
		}

		//request immutable ids so the returned draft id survives the draft->Sent Items move and is safe
		//to store as the next thread anchor
		try {
			$message = $this->graph->users()->byUserId( $from )->messages()->byMessageId( $messageId )->createReply()
			                       ->post( new CreateReplyPostRequestBody(), new CreateReplyRequestBuilderPostRequestConfiguration( headers: [ 'Prefer' => self::IMMUTABLE_ID ] ) )
			                       ->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to create reply draft', $e );
		}

		if( !$message instanceof Message ) {
			throw new serviceException( 'Graph did not return the reply draft it created', 500 );
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
	 * @return \Microsoft\Graph\Generated\Models\Message The sent message; read ->getId() (the next anchor) and ->getConversationId()
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function replyIntoThread( string $from, string $anchorMessageId, string|array $to, string|array $bcc, string $htmlBody, string|array $cc = [] ): Message {
		if( $from=='' ) {
			$from = $this->config->fromAddress;
		}

		//1) create the reply draft in the same conversation as the anchor
		$draft   = $this->createReplyDraft( $anchorMessageId, $from );
		$draftId = (string)$draft->getId();

		//2) set recipients + body on the draft (subject and threading headers are inherited from the anchor)
		$patch = new Message();
		$patch->setToRecipients( $this->buildRecipients( $to ) );
		$patch->setCcRecipients( $this->buildRecipients( $cc ) );
		$patch->setBccRecipients( $this->buildRecipients( $bcc ) );
		$patch->setBody( $this->htmlBody( $htmlBody ) );

		$draftRequest = $this->graph->users()->byUserId( $from )->messages()->byMessageId( $draftId );

		try {
			$message = $draftRequest->patch( $patch, new MessageItemRequestBuilderPatchRequestConfiguration( headers: [ 'Prefer' => self::IMMUTABLE_ID ] ) )->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to update reply draft', $e );
		}

		if( !$message instanceof Message ) {
			throw new serviceException( 'Graph did not return the reply draft it updated', 500 );
		}

		//3) send the draft
		try {
			$draftRequest->send()->post()->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to send reply', $e );
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
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getLatestMessageInConversation( string $from, string $conversationId ): ?Message {
		if( $from=='' ) {
			$from = $this->config->fromAddress;
		}

		//escape single quotes for the OData string literal; the SDK url-encodes the query values
		$escapedConversationId = str_replace( "'", "''", $conversationId );

		$requestConfiguration = new MessagesRequestBuilderGetRequestConfiguration(
			headers        : [ 'Prefer' => self::IMMUTABLE_ID ],
			queryParameters: MessagesRequestBuilderGetRequestConfiguration::createQueryParameters(
				filter : "conversationId eq '" . $escapedConversationId . "'",
				orderby: [ 'receivedDateTime desc' ],
				select : [ 'id', 'conversationId' ],
				top    : 1
			)
		);

		try {
			$messages = $this->graph->users()->byUserId( $from )->messages()->get( $requestConfiguration )->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to look up conversation', $e );
		}

		return $messages?->getValue()[ 0 ] ?? null;
	}


	/**
	 * @return \Microsoft\Graph\Generated\Models\Recipient[] One recipient for each address. Empty values are skipped.
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
			$recipients[] = $this->recipient( $emailAddress );
		}

		return $recipients;
	}


	private function recipient( string $address ): Recipient {
		$emailAddress = new EmailAddress();
		$emailAddress->setAddress( $address );

		$recipient = new Recipient();
		$recipient->setEmailAddress( $emailAddress );

		return $recipient;
	}


	private function htmlBody( string $content ): ItemBody {
		$body = new ItemBody();
		$body->setContentType( new BodyType( BodyType::HTML ) );
		$body->setContent( $content );

		return $body;
	}


	/**
	 * @param string                                              $emailAddress
	 * @param \andrewsauder\microsoftServices\mail\headers\prefer $preferHeader [optional] Defaults to html
	 *
	 * @return \Microsoft\Graph\Generated\Models\Message[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getAllMessages( string $emailAddress, prefer $preferHeader = prefer::HTML ): array {
		try {
			return $this->collectAll(
				$this->graph->users()->byUserId( $emailAddress )->messages()
				            ->get( new MessagesRequestBuilderGetRequestConfiguration( headers: $this->listMessagesHeaders( $preferHeader ) ) )
				            ->wait()
			);
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to get messages', $e );
		}
	}


	/**
	 * @param string                                              $emailAddress
	 * @param string                                              $mailFolderId
	 * @param \andrewsauder\microsoftServices\mail\headers\prefer $preferHeader [optional] Defaults to html
	 *
	 * @return \Microsoft\Graph\Generated\Models\Message[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getMessagesInFolder( string $emailAddress, string $mailFolderId, prefer $preferHeader = prefer::HTML ): array {
		try {
			return $this->collectAll(
				$this->graph->users()->byUserId( $emailAddress )->mailFolders()->byMailFolderId( $mailFolderId )->messages()
				            ->get( new FolderMessagesRequestBuilderGetRequestConfiguration( headers: $this->listMessagesHeaders( $preferHeader ) ) )
				            ->wait()
			);
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to get messages', $e );
		}
	}


	/** @return array<string, string> */
	private function listMessagesHeaders( prefer $preferHeader ): array {
		return [ 'Prefer' => 'outlook.body-content-type="' . $preferHeader->value . '", ' . self::IMMUTABLE_ID ];
	}


	/**
	 * @param string $emailAddress
	 * @param bool   $includeHiddenFolders
	 *
	 * @return \Microsoft\Graph\Generated\Models\MailFolder[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getFolders( string $emailAddress, bool $includeHiddenFolders = false ): array {
		$requestConfiguration = new MailFoldersRequestBuilderGetRequestConfiguration(
			queryParameters: MailFoldersRequestBuilderGetRequestConfiguration::createQueryParameters(
				includeHiddenFolders: $includeHiddenFolders ? 'true' : null
			)
		);

		try {
			return $this->collectAll( $this->graph->users()->byUserId( $emailAddress )->mailFolders()->get( $requestConfiguration )->wait() );
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to get mail folders', $e );
		}
	}


	/**
	 * The attachments of a message. A file attachment is a \Microsoft\Graph\Generated\Models\FileAttachment, and its
	 * getContentBytes() stream holds the base64 text of the file.
	 *
	 * @param string $emailAddress
	 * @param string $messageId
	 *
	 * @return \Microsoft\Graph\Generated\Models\Attachment[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getAttachments( string $emailAddress, string $messageId ): array {
		try {
			return $this->collectAll( $this->graph->users()->byUserId( $emailAddress )->messages()->byMessageId( $messageId )->attachments()->get()->wait() );
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to get attachments', $e );
		}
	}


	/**
	 * @param string $emailAddress
	 * @param string $messageId
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException when Graph does not delete the message
	 */
	public function deleteMessage( string $emailAddress, string $messageId ): void {
		try {
			$this->graph->users()->byUserId( $emailAddress )->messages()->byMessageId( $messageId )->delete()->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to delete email message', $e );
		}
	}


	/**
	 * @param string $emailAddress
	 * @param string $messageId
	 * @param string $mailFolderId Folder id, or a well-known folder name such as "archive" or "deleteditems"
	 *
	 * @return \Microsoft\Graph\Generated\Models\Message The message in its new folder. It has a new id unless the id is immutable.
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function moveMessage( string $emailAddress, string $messageId, string $mailFolderId ): Message {
		$requestBody = new MovePostRequestBody();
		$requestBody->setDestinationId( $mailFolderId );

		try {
			$message = $this->graph->users()->byUserId( $emailAddress )->messages()->byMessageId( $messageId )->move()->post( $requestBody )->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to move email message', $e );
		}

		if( !$message instanceof Message ) {
			throw new serviceException( 'Graph did not return the moved message', 500 );
		}

		return $message;
	}

}
