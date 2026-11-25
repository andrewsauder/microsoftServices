<?php

namespace andrewsauder\microsoftServices\mail;

use andrewsauder\microsoftServices\exceptions\serviceException;


/**
 * One file that a message attaches. mail::addAttachment() creates it; mail::createDraft() and mail::send()
 * read it. The file content is read only when a request needs it, so a queued attachment costs no memory.
 *
 * Graph offers two ways to attach a file to a draft message:
 * - a single POST of the base64 content, for a file under 3 MB
 * - an upload session, for a file from 3 MB to 150 MB, which takes the raw bytes in ranges of at most 4 MB
 *
 * Every Graph request must stay under 4 MB, so neither route puts more than about 4 MB in one request.
 *
 * @see https://learn.microsoft.com/graph/outlook-large-attachments
 */
final class attachment {

	/**
	 * The largest file this library POSTs in one request. Base64 makes the content about 4/3 larger,
	 * so a 3,000,000 byte file is 4,000,000 bytes once encoded - under the Graph 4 MB request limit.
	 */
	public const INLINE_MAX_BYTES = 3_000_000;

	/** The largest file Graph accepts through an upload session */
	public const UPLOAD_SESSION_MAX_BYTES = 150 * 1024 * 1024;

	/** Bytes sent in each upload session PUT. Graph asks for at most 4 MB; 10 x 320 KiB stays under it. */
	public const CHUNK_BYTES = 10 * 320 * 1024;

	public readonly string $path;

	public readonly string $name;

	public readonly string $contentType;

	public readonly int $size;


	public function __construct( string $path, string $name, string $contentType, int $size ) {
		$this->path        = $path;
		$this->name        = $name;
		$this->contentType = $contentType;
		$this->size        = $size;
	}


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException 500 when the file does not exist or cannot be read, 400 when it is larger than Graph accepts
	 */
	public static function fromFile( string $filePath ): self {
		if( !is_file( $filePath ) || !is_readable( $filePath ) ) {
			throw new serviceException( 'File does not exist', 500 );
		}

		$size = filesize( $filePath );
		if( $size===false ) {
			throw new serviceException( 'Unable to read the size of ' . basename( $filePath ), 500 );
		}
		if( $size>self::UPLOAD_SESSION_MAX_BYTES ) {
			throw new serviceException( basename( $filePath ) . ' is larger than the 150 MB Graph accepts for one attachment', 400 );
		}

		$contentType = mime_content_type( $filePath );

		return new self( $filePath, basename( $filePath ), $contentType===false ? 'application/octet-stream' : $contentType, $size );
	}


	/** True when the file is too large for one POST and goes through an upload session */
	public function usesUploadSession(): bool {
		return $this->size>self::INLINE_MAX_BYTES;
	}


	/**
	 * The fileAttachment resource with the base64 content, for a POST to a message's attachments or for
	 * the attachments list of a sendMail request
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException 500 when the file cannot be read
	 */
	public function fileAttachmentBody(): array {
		$contents = @file_get_contents( $this->path );
		if( $contents===false ) {
			throw new serviceException( 'Unable to read the attachment ' . $this->name, 500 );
		}

		return [
			'@odata.type'  => '#microsoft.graph.fileAttachment',
			'name'         => $this->name,
			'contentType'  => $this->contentType,
			'contentBytes' => base64_encode( $contents )
		];
	}


	/** The request body of createUploadSession */
	public function uploadSessionBody(): array {
		return [
			'AttachmentItem' => [
				'attachmentType' => 'file',
				'name'           => $this->name,
				'size'           => $this->size,
				'contentType'    => $this->contentType
			]
		];
	}


	/**
	 * The byte ranges an upload session PUTs, in order. Each range is inclusive at both ends, the way the
	 * Content-Range header states it.
	 *
	 * @return array<int, array{0: int, 1: int}> [ first byte, last byte ] pairs
	 */
	public static function byteRanges( int $size, int $chunkBytes = self::CHUNK_BYTES ): array {
		if( $size<=0 || $chunkBytes<=0 ) {
			return [];
		}

		$ranges = [];
		for( $start = 0; $start<$size; $start += $chunkBytes ) {
			$ranges[] = [ $start, min( $start + $chunkBytes, $size ) - 1 ];
		}

		return $ranges;
	}


	/** The Content-Range header value for one range of a file of $size bytes */
	public static function contentRange( int $first, int $last, int $size ): string {
		return 'bytes ' . $first . '-' . $last . '/' . $size;
	}

}
