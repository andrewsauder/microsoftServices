<?php

namespace andrewsauder\microsoftServices;

use andrewsauder\microsoftServices\exceptions\serviceException;
use GuzzleHttp\Exception\GuzzleException;
use Microsoft\Graph\Exception\GraphException;


class mail {

	private \andrewsauder\microsoftServices\config $config;
	private ?string $userAccessToken = null;
	private array $attachments = [];


	/**
	 * @param  string|null  $userAccessToken  Provide user token. If config.onBehalfOfFlow is enabled, the provided token will be exchanged for an access token for this API. If config.onBehalfOfFlow is not enabled, this token will be used
	 *                                        for the request. If no token is provided, an application token will be generated
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function __construct( \andrewsauder\microsoftServices\config $config, ?string $userAccessToken=null ) {
		$config->validateForMail();
		$this->config = $config;
		$this->userAccessToken = $userAccessToken;
	}




	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function addAttachment( string $filePath ) {
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
	}


	/**
	 * @param  string|string[]  $to
	 * @param  string           $subject
	 * @param  string           $content
	 * @param  string           $from
	 *
	 * @return \Microsoft\Graph\Http\GraphResponse
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function send( string|array $to, string $subject, string $content, string $from = '' ) : \Microsoft\Graph\Http\GraphResponse {
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

		if( $from == '' ) {
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

		if( count( $this->attachments ) > 0 ) {
			$mailBody[ 'Message' ][ 'attachments' ] = $this->attachments;
		}


		try {
			$response = $graph->createRequest( 'POST', '/users/' . $from . '/sendMail' )->attachBody( $mailBody )->execute();
		}
		catch(GraphException $e) {
			error_log($e);
			throw new serviceException( 'Failed to send email: '.$e->getMessage(), 500, $e );
		}
		catch(GuzzleException $e){
			error_log($e);
			throw new serviceException( 'Failed to send email: '.$e->getMessage(), $e->getCode(), $e );
		}
		return $response;
	}


	/**
	 * @return string
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	private function getMicrosoftAccessToken() : string {
		$microsoftAuth = new auth( $this->config );

		if( !$this->config->onBehalfOfFlow && isset( $this->userAccessToken ) ) {
			return $this->userAccessToken;
		}
		//get OBO user access token
		if( $this->config->onBehalfOfFlow && isset( $this->userAccessToken ) ) {
			return (string) $microsoftAuth->getAccessToken( $this->userAccessToken );
		}
		//get access token
		else {
			return $microsoftAuth->getApplicationAccessToken();
		}
	}

}