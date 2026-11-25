<?php
namespace andrewsauder\microsoftServices;

use andrewsauder\microsoftServices\components\envelope;
use andrewsauder\microsoftServices\exceptions\serviceException;
use GuzzleHttp\Psr7\Utils;
use Microsoft\Graph\Generated\Drives\Item\Items\Item\CreateUploadSession\CreateUploadSessionPostRequestBody;
use Microsoft\Graph\Generated\Drives\Item\Items\Item\DriveItemItemRequestBuilder;
use Microsoft\Graph\Generated\Drives\Item\Items\Item\DriveItemItemRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\Generated\Drives\Item\Items\Item\Preview\PreviewPostRequestBody;
use Microsoft\Graph\Generated\Models\DriveItem;
use Microsoft\Graph\Generated\Models\DriveItemUploadableProperties;
use Microsoft\Graph\Generated\Models\FileSystemInfo;
use Microsoft\Graph\Generated\Models\Folder;
use Microsoft\Graph\Generated\Models\ItemPreviewInfo;
use Microsoft\Graph\Generated\Models\ItemReference;
use Microsoft\Graph\Generated\Models\Thumbnail;

class files extends \andrewsauder\microsoftServices\components\service {

	/** Files up to this size go up in one PUT; larger files go through an upload session */
	public const SIMPLE_UPLOAD_MAX_BYTES = 4194304;

	public string $rootBasePath = '';


	/**
	 * @param \andrewsauder\microsoftServices\config $config
	 * @param string|null                            $userAccessToken Provide user token. If config.onBehalfOfFlow is enabled, the provided token will be exchanged for an access token for this API. If config.onBehalfOfFlow is not
	 *                                                                enabled, this token will be used for the request. If no token is provided, an application token will be generated
	 * @param string                                 $rootBasePath
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function __construct( \andrewsauder\microsoftServices\config $config, ?string $userAccessToken = null, string $rootBasePath = '' ) {
		parent::__construct( $config, $userAccessToken );

		$config->validateForFiles();

		if( strlen( trim( $rootBasePath, ' \\/' ) )>0 ) {
			$this->rootBasePath = trim( $rootBasePath, ' \\/' ) . '/';
		}
	}


	/**
	 * The Graph URL of a drive item by its path from the drive root, for example
	 * https://graph.microsoft.com/v1.0/drives/{driveId}/root:/2021-0001/Building%201:
	 *
	 * Each path segment is url-encoded on its own, so the ':' and '/' that Graph reads stay as they are. An empty
	 * path is the drive root.
	 */
	private function driveItemPathUrl( string $path ): string {
		$segments = array_values( array_filter( explode( '/', $path ), static fn( string $segment ) => $segment!=='' ) );

		$url = rtrim( $this->graph->getRequestAdapter()->getBaseUrl(), '/' ) . '/drives/' . rawurlencode( $this->config->driveId ) . '/root';
		if( count( $segments )>0 ) {
			$url .= ':/' . implode( '/', array_map( 'rawurlencode', $segments ) ) . ':';
		}

		return $url;
	}


	/** The request builder of a drive item by its id */
	private function driveItemById( string $itemId ): DriveItemItemRequestBuilder {
		return $this->graph->drives()->byDriveId( $this->config->driveId )->items()->byDriveItemId( $itemId );
	}


	/** The request builder of a drive item by its path from the drive root */
	private function driveItemByPath( string $path ): DriveItemItemRequestBuilder {
		return $this->driveItemById( 'root' )->withUrl( $this->driveItemPathUrl( $path ) );
	}


	/**
	 * @param string[] $microsoftPathParts Ex: [ '2021-0001', 'Building 1', 'Inspections' ] will turn into {root}/2021-0001/Building 1/Inspections
	 *
	 * @return \Microsoft\Graph\Generated\Models\DriveItem[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function list( array $microsoftPathParts = [], bool $recursive = true ): array {
		try {
			return $this->getMicrosoftDriveItems( implode( '/', $microsoftPathParts ), $recursive );
		}
		catch( serviceException $e ) {
			//a folder that does not exist has no files
			if( $e->getCode()==404 ) {
				return [];
			}
			throw $e;
		}
	}


	/**
	 * @param string $id
	 *
	 * @return \Microsoft\Graph\Generated\Models\DriveItem[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function listById( string $id, bool $recursive = true ): array {
		try {
			return $this->getMicrosoftDriveItemsById( $id, $recursive );
		}
		catch( serviceException $e ) {
			//a folder that does not exist has no files
			if( $e->getCode()==404 ) {
				return [];
			}
			throw $e;
		}
	}


	/**
	 * @param string[] $microsoftPathParts Ex: [ '2021-0001', 'Building 1', 'Inspections' ] will turn into {root}/2021-0001/Building 1/Inspections. An empty array is the drive root.
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getDriveItem( array $microsoftPathParts ): DriveItem {
		$path = count( $microsoftPathParts )>0 ? $this->rootBasePath . implode( '/', $microsoftPathParts ) : '';

		try {
			$driveItem = $this->driveItemByPath( $path )->get()->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to get file', $e );
		}

		return $this->requireDriveItem( $driveItem, 'File not found' );
	}


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getDriveItemById( string $itemId ): DriveItem {
		try {
			$driveItem = $this->driveItemById( $itemId )->get()->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to get file', $e );
		}

		return $this->requireDriveItem( $driveItem, 'File not found' );
	}


	/**
	 * @param int|null $pageNumber The page of the document to start at
	 * @param int|null $zoom       The zoom level to start at
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getDriveItemPreviewById( string $itemId, ?int $pageNumber = null, ?int $zoom = null ): ItemPreviewInfo {
		$requestBody = new PreviewPostRequestBody();
		if( !empty( $pageNumber ) ) {
			$requestBody->setPage( (string)$pageNumber );
		}
		if( !empty( $zoom ) ) {
			$requestBody->setZoom( (float)$zoom );
		}

		try {
			$previewInfo = $this->driveItemById( $itemId )->preview()->post( $requestBody )->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to get file preview', $e );
		}

		if( !$previewInfo instanceof ItemPreviewInfo ) {
			throw new serviceException( 'Graph did not return a preview', 500 );
		}

		return $previewInfo;
	}


	/**
	 * @param string $size small, medium, large, or source
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException 400 for another size, 404 when the thumbnail does not exist
	 */
	public function getDriveItemThumbnailById( string $itemId, int $thumbnailId = 0, string $size = 'medium' ): Thumbnail {
		if( !in_array( $size, [ 'small', 'medium', 'large', 'source' ], true ) ) {
			throw new serviceException( 'Thumbnail size must be small, medium, large, or source', 400 );
		}

		try {
			$thumbnailSet = $this->driveItemById( $itemId )->thumbnails()->byThumbnailSetId( (string)$thumbnailId )->get()->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to get file thumbnail', $e );
		}

		$thumbnail = match ( $size ) {
			'small'  => $thumbnailSet?->getSmall(),
			'medium' => $thumbnailSet?->getMedium(),
			'large'  => $thumbnailSet?->getLarge(),
			'source' => $thumbnailSet?->getSource(),
		};

		if( !$thumbnail instanceof Thumbnail ) {
			throw new serviceException( 'Thumbnail not found', 404 );
		}

		return $thumbnail;
	}


	/**
	 * @param string $itemId  Microsoft file id
	 * @param string $tmpPath Path to download the file into
	 *
	 * @return string File path name to tmp local copy of file
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function downloadDriveItemById( string $itemId, string $tmpPath = '' ): string {
		try {
			$driveItem = $this->driveItemById( $itemId )->get( new DriveItemItemRequestBuilderGetRequestConfiguration(
				queryParameters: DriveItemItemRequestBuilderGetRequestConfiguration::createQueryParameters( select: [ 'id', 'name' ] )
			) )->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to get file', $e );
		}

		$driveItem = $this->requireDriveItem( $driveItem, 'File not found' );

		//download the file into {tmpPath}/{itemId}/{name}
		$path = rtrim( $tmpPath, '/' ) . '/' . $itemId;
		if( !file_exists( $tmpPath ) ) {
			mkdir( $tmpPath );
		}
		if( !file_exists( $path ) ) {
			mkdir( $path );
		}
		$filePathName = $path . '/' . $driveItem->getName();

		if( file_exists( $filePathName ) ) {
			return $filePathName;
		}

		try {
			$content = $this->driveItemById( $itemId )->content()->get()->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to download file', $e );
		}

		$handle = @fopen( $filePathName, 'wb' );
		if( $handle===false ) {
			throw new serviceException( 'Unable to write ' . $filePathName, 500 );
		}
		try {
			while( $content!==null && !$content->eof() ) {
				fwrite( $handle, $content->read( 1048576 ) );
			}
		}
		finally {
			fclose( $handle );
		}

		return $filePathName;
	}


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function moveItem( string $itemIdToMove, string $newParentDirItemId ): DriveItem {
		return $this->update( $itemIdToMove, [ 'parentReference' => [ 'id' => $newParentDirItemId ] ] );
	}


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function renameItem( string $itemIdToRename, string $newName ): DriveItem {
		return $this->update( $itemIdToRename, [ 'name' => $newName ] );
	}


	/**
	 * Apply a partial update to a drive item in a single Graph PATCH. Supports renaming and/or moving in one
	 * request. Accepts a raw driveItem patch body, eg [ 'name' => 'new-name.txt' ] and/or
	 * [ 'parentReference' => [ 'id' => '{destinationFolderItemId}' ] ]. Other keys are sent as they are.
	 *
	 * @param string               $itemId         Microsoft drive item id
	 * @param array<string, mixed> $driveItemPatch Partial driveItem body
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function update( string $itemId, array $driveItemPatch ): DriveItem {
		//nothing to change - return the item as-is rather than issuing an empty PATCH (Graph rejects it as 400)
		if( count( $driveItemPatch )===0 ) {
			return $this->getDriveItemById( $itemId );
		}

		try {
			$driveItem = $this->driveItemById( $itemId )->patch( self::driveItemPatch( $driveItemPatch ) )->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to update file', $e );
		}

		return $this->requireDriveItem( $driveItem, 'Graph did not return the updated file' );
	}


	/**
	 * The DriveItem body of a PATCH: name and parentReference as typed properties, every other key as it is
	 *
	 * @param array<string, mixed> $driveItemPatch
	 */
	public static function driveItemPatch( array $driveItemPatch ): DriveItem {
		$driveItem = new DriveItem();

		if( array_key_exists( 'name', $driveItemPatch ) ) {
			$driveItem->setName( $driveItemPatch[ 'name' ] );
			unset( $driveItemPatch[ 'name' ] );
		}

		if( isset( $driveItemPatch[ 'parentReference' ] ) && is_array( $driveItemPatch[ 'parentReference' ] ) ) {
			$parentReference = new ItemReference();
			if( isset( $driveItemPatch[ 'parentReference' ][ 'id' ] ) ) {
				$parentReference->setId( $driveItemPatch[ 'parentReference' ][ 'id' ] );
			}
			if( isset( $driveItemPatch[ 'parentReference' ][ 'driveId' ] ) ) {
				$parentReference->setDriveId( $driveItemPatch[ 'parentReference' ][ 'driveId' ] );
			}
			$driveItem->setParentReference( $parentReference );
			unset( $driveItemPatch[ 'parentReference' ] );
		}

		if( count( $driveItemPatch )>0 ) {
			$driveItem->setAdditionalData( $driveItemPatch );
		}

		return $driveItem;
	}


	/**
	 * @param string   $base64EncodedContent
	 * @param string   $contentType
	 * @param string   $fileName
	 * @param string[] $uploadPathParts  Ex: [ '2021-0001', 'Building 1', 'Inspections' ] will turn into {root}/2021-0001/Building 1/Inspections
	 * @param string   $conflictBehavior The conflict resolution behavior for actions that create a new item. You can use the values fail, replace, or rename. The default for PUT is replace.
	 */
	public function uploadBase64EncodedContent( string $base64EncodedContent, string $contentType, string $fileName, array $uploadPathParts = [], string $conflictBehavior = 'replace' ): \andrewsauder\microsoftServices\components\upload {
		$tempFileName = tempnam( sys_get_temp_dir(), 'MicrosoftServicesFile' );
		if( $tempFileName===false || file_put_contents( $tempFileName, base64_decode( $base64EncodedContent ) )===false ) {
			$response           = new \andrewsauder\microsoftServices\components\upload();
			$response->errors[] = new envelope( 500, true, 'Unable to write to tmp directory at ' . sys_get_temp_dir() );
			return $response;
		}

		try {
			return $this->upload( $tempFileName, $fileName, $uploadPathParts, $conflictBehavior );
		}
		finally {
			@unlink( $tempFileName );
		}
	}


	/**
	 * Upload a file. A file up to 4 MB goes up in one PUT. A larger file goes through an upload session in ranges.
	 * This method does not throw: a failure is in the errors of the response.
	 *
	 * @param string   $serverFullFilePath
	 * @param string   $fileName
	 * @param string[] $uploadPathParts  Ex: [ '2021-0001', 'Building 1', 'Inspections' ] will turn into {root}/2021-0001/Building 1/Inspections
	 * @param string   $conflictBehavior The conflict resolution behavior for actions that create a new item. You can use the values fail, replace, or rename. The default for PUT is replace.
	 */
	public function upload( string $serverFullFilePath, string $fileName, array $uploadPathParts = [], string $conflictBehavior = 'replace' ): \andrewsauder\microsoftServices\components\upload {
		$response = new \andrewsauder\microsoftServices\components\upload();

		$filePath = $this->rootBasePath . implode( '/', $uploadPathParts ) . '/' . $fileName;
		$fileUrl  = $this->driveItemPathUrl( $filePath );

		try {
			$fileSize = filesize( $serverFullFilePath );
			if( $fileSize===false ) {
				throw new serviceException( 'Unable to read ' . $fileName, 500 );
			}

			if( $fileSize<=self::SIMPLE_UPLOAD_MAX_BYTES ) {
				$handle = @fopen( $serverFullFilePath, 'rb' );
				if( $handle===false ) {
					throw new serviceException( 'Unable to read ' . $fileName, 500 );
				}
				$driveItem = $this->driveItemById( 'root' )->content()
				                  ->withUrl( $fileUrl . '/content?@microsoft.graph.conflictBehavior=' . rawurlencode( $conflictBehavior ) )
				                  ->put( Utils::streamFor( $handle ) )
				                  ->wait();
			}
			else {
				$uploadUrl = $this->createUploadSession( $fileUrl, $fileName, $conflictBehavior );

				$status = $this->uploadFileInRanges( $uploadUrl, $serverFullFilePath, $fileSize, $fileName );
				if( $status!==200 && $status!==201 ) {
					throw new serviceException( 'Graph did not complete the upload of ' . $fileName . ' (HTTP ' . $status . ')', 500 );
				}

				$driveItem = $this->driveItemByPath( $filePath )->get()->wait();
			}

			$response->files[] = $this->requireDriveItem( $driveItem, 'Graph did not return the uploaded file' );
		}
		catch( \Throwable $e ) {
			$serviceException   = $this->toServiceException( $fileName . ' did not upload', $e );
			$response->errors[] = new envelope( $serviceException->getCode(), true, $serviceException->getMessage() );
		}

		return $response;
	}


	/**
	 * @return string The upload URL of the new session
	 * @throws \Exception
	 */
	private function createUploadSession( string $fileUrl, string $fileName, string $conflictBehavior ): string {
		$item = new DriveItemUploadableProperties();
		$item->setName( $fileName );
		$item->setDescription( '' );
		$item->setFileSystemInfo( new FileSystemInfo() );
		$item->setAdditionalData( [ '@microsoft.graph.conflictBehavior' => $conflictBehavior ] );

		$requestBody = new CreateUploadSessionPostRequestBody();
		$requestBody->setItem( $item );

		$uploadSession = $this->driveItemById( 'root' )->createUploadSession()
		                      ->withUrl( $fileUrl . '/createUploadSession' )
		                      ->post( $requestBody )
		                      ->wait();

		$uploadUrl = (string)$uploadSession?->getUploadUrl();
		if( $uploadUrl==='' ) {
			throw new serviceException( 'Graph did not return an upload URL for ' . $fileName, 500 );
		}

		return $uploadUrl;
	}


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function delete( string $itemId ): void {
		try {
			$this->driveItemById( $itemId )->delete()->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to delete file', $e );
		}
	}


	/**
	 * @param string   $folderName
	 * @param string[] $basePathParts Ex: [ '2021-0001', 'Building 1', 'Inspections' ] will turn into {root}/2021-0001/Building 1/Inspections
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException 409 when the folder already exists
	 */
	public function createFolder( string $folderName, array $basePathParts = [] ): DriveItem {
		$folder = new DriveItem();
		$folder->setName( $folderName );
		$folder->setFolder( new Folder() );
		$folder->setAdditionalData( [ '@microsoft.graph.conflictBehavior' => 'fail' ] );

		$parentPath = count( $basePathParts )>0 ? $this->rootBasePath . implode( '/', $basePathParts ) : '';

		try {
			$driveItem = $this->driveItemById( 'root' )->children()
			                  ->withUrl( $this->driveItemPathUrl( $parentPath ) . '/children' )
			                  ->post( $folder )
			                  ->wait();
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( $this->rootBasePath . implode( '/', $basePathParts ) . '/' . $folderName . ' not created', $e );
		}

		return $this->requireDriveItem( $driveItem, 'Graph did not return the new folder' );
	}


	/**
	 * @param string $path Path below rootBasePath
	 *
	 * @return \Microsoft\Graph\Generated\Models\DriveItem[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	private function getMicrosoftDriveItems( string $path, bool $recursive = true ): array {
		try {
			/** @var \Microsoft\Graph\Generated\Models\DriveItem[] $driveItems */
			$driveItems = $this->collectAll(
				$this->driveItemById( 'root' )->children()
				     ->withUrl( $this->driveItemPathUrl( $this->rootBasePath . $path ) . '/children' )
				     ->get()
				     ->wait()
			);
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to list folder', $e );
		}

		if( $recursive ) {
			foreach( $driveItems as $driveItem ) {
				if( $driveItem->getFolder()!==null ) {
					$driveItem->setChildren( $this->getMicrosoftDriveItems( $path . '/' . $driveItem->getName() ) );
				}
			}
		}

		return $driveItems;
	}


	/**
	 * @return \Microsoft\Graph\Generated\Models\DriveItem[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	private function getMicrosoftDriveItemsById( string $id, bool $recursive = true ): array {
		try {
			/** @var \Microsoft\Graph\Generated\Models\DriveItem[] $driveItems */
			$driveItems = $this->collectAll( $this->driveItemById( $id )->children()->get()->wait() );
		}
		catch( \Throwable $e ) {
			throw $this->toServiceException( 'Failed to list folder', $e );
		}

		if( $recursive ) {
			foreach( $driveItems as $driveItem ) {
				if( $driveItem->getFolder()!==null ) {
					$driveItem->setChildren( $this->getMicrosoftDriveItemsById( (string)$driveItem->getId() ) );
				}
			}
		}

		return $driveItems;
	}


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException 404 when Graph returned no drive item
	 */
	private function requireDriveItem( mixed $driveItem, string $message ): DriveItem {
		if( !$driveItem instanceof DriveItem ) {
			throw new serviceException( $message, 404 );
		}

		return $driveItem;
	}


		public static function sanitizeFileFolderName(string $name) {
		// Step 1: Remove forbidden characters
		$forbiddenChars = ['~', '#', '%', '&', '*', '{', '}', '\\', ':', '<', '>', '?', '/', '+', '|', '"'];
		$name = str_replace($forbiddenChars, '', $name);

		// Step 2: Replace multiple consecutive periods with a single one
		$name = preg_replace('/\.{2,}/', '.', $name);

		// Step 3a: Remove leading or trailing periods and spaces
		$name = trim($name, '. ');

		// Step 4: List of disallowed suffixes (case-insensitive)
		$disallowedSuffixes = [
			'.files', '_files', '-Dateien', '_fichiers', '_bestanden', '_file', '_archivos',
			'-filer', '_tiedostot', '_pliki', '_soubory', '_elemei', '_ficheiros', '_arquivos',
			'_dosyalar', '_datoteke', '_fitxers', '_failid', '_fails', '_bylos', '_fajlovi', '_fitxategiak'
		];

		// Step 5: Check and remove disallowed suffixes
		foreach ($disallowedSuffixes as $suffix) {
			if (str_ends_with(strtolower($name), strtolower($suffix))) {
				$name = substr($name, 0, -strlen($suffix));
				break;
			}
		}

		// Step 6: If name is empty after sanitization, assign a default
		if ($name === '') {
			$name = bin2hex(openssl_random_pseudo_bytes(10));;
		}

		return $name;
	}

}
