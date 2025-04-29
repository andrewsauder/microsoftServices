<?php
namespace andrewsauder\microsoftServices;

use andrewsauder\microsoftServices\exceptions\serviceException;
use GuzzleHttp\Exception\ClientException;
use GuzzleHttp\Exception\GuzzleException;
use JetBrains\PhpStorm\Deprecated;
use Microsoft\Graph\Exception\GraphException;

class files extends \andrewsauder\microsoftServices\components\service {

	public string                                  $rootBasePath    = '';

	/**
	 * @param  \andrewsauder\microsoftServices\config  $config
	 * @param  string|null                             $userAccessToken  Provide user token. If config.onBehalfOfFlow is enabled, the provided token will be exchanged for an access token for this API. If config.onBehalfOfFlow is not
	 *                                                                   enabled, this token will be used for the request. If no token is provided, an application token will be generated
	 * @param  string                                  $rootBasePath
	 *
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function __construct( \andrewsauder\microsoftServices\config $config, ?string $userAccessToken = null, string $rootBasePath = '' ) {
		parent::__construct( $config, $userAccessToken );

		$config->validateForFiles();

		if( strlen( trim( $rootBasePath, ' \\/' ) ) > 0 ) {
			$this->rootBasePath = trim( $rootBasePath, ' \\/' ) . '/';
		}
	}


	/**
	 * @param  string[]  $microsoftPathParts  Ex: [ '2021-0001', 'Building 1', 'Inspections' ] will turn into {root}/2021-0001/Building 1/Inspections
	 *
	 * @return \Microsoft\Graph\Model\DriveItem[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function list( array $microsoftPathParts=[], bool $recursive=true ) : array {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//get file list
		try {
			$driveItems = $this->getMicrosoftDriveItems( $accessToken, implode( '/', $microsoftPathParts ), $recursive );
		}
		catch( serviceException $e ) {
			//if the error is that the folder doesn't exist, try to create it
			if( $e->getCode() == 404 ) {
				//generate the folders recursively
				//$newDriveItems = $this->createMicrosoftDirectories( $accessToken, $microsoftPathParts );
				//try to get the files again
				//$driveItems = $this->getMicrosoftDriveItems( $accessToken, implode( '/', $microsoftPathParts ) );
				$driveItems = [];
			}
			else {
				throw new serviceException( $e->getMessage(), $e->getCode(), $e );
			}
		}

		return $driveItems;
	}

	/**
	 * @param  string $id
	 *
	 * @return \Microsoft\Graph\Model\DriveItem[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function listById( string $id, bool $recursive=true ) : array {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//get file list
		try {
			$driveItems = $this->getMicrosoftDriveItemsById( $accessToken, $id, $recursive );
		}
		catch( serviceException $e ) {
			//if the error is that the folder doesn't exist, try to create it
			if( $e->getCode() == 404 ) {
				//generate the folders recursively
				//$newDriveItems = $this->createMicrosoftDirectories( $accessToken, $microsoftPathParts );
				//try to get the files again
				//$driveItems = $this->getMicrosoftDriveItems( $accessToken, implode( '/', $microsoftPathParts ) );
				$driveItems = [];
			}
			else {
				throw new serviceException( $e->getMessage(), $e->getCode(), $e );
			}
		}

		return $driveItems;
	}


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getDriveItem( array $microsoftPathParts ) : \Microsoft\Graph\Model\DriveItem {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//get file list
		try {
			$graph = new \Microsoft\Graph\Graph();
			$graph->setAccessToken( $accessToken );

			/** @var \Microsoft\Graph\Model\DriveItem $driveItem */
			$driveItem = $graph->createRequest( "GET", "/drives/" . $this->config->driveId . '/root:/' . $this->rootBasePath . implode( '/', $microsoftPathParts ) )->setReturnType( \Microsoft\Graph\Model\DriveItem::class )->execute();

			return $driveItem;
		}
		catch( ClientException $e ) {
			throw new serviceException( 'File not found', $e->getCode(), $e );
		}
		catch( GuzzleException $e ) {
			throw new serviceException( 'Error getting files from Microsoft', 500, $e );
		}
		catch( GraphException $e ) {
			throw new serviceException( $e->getMessage(), 500, $e );
		}
	}

	#[Deprecated('Renamed method', '%class%->getDriveItem()')]
	public function getFile( array $microsoftPathParts ) : \Microsoft\Graph\Model\DriveItem {
		return $this->getDriveItem( $microsoftPathParts );
	}


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function getDriveItemById( string $itemId ) : \Microsoft\Graph\Model\DriveItem {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//get file list
		try {
			$graph = new \Microsoft\Graph\Graph();
			$graph->setAccessToken( $accessToken );

			/** @var \Microsoft\Graph\Model\DriveItem $driveItem */
			$driveItem = $graph->createRequest( "GET", "/drives/" . $this->config->driveId . '/root:/' . $itemId )->setReturnType( \Microsoft\Graph\Model\DriveItem::class )->execute();

			return $driveItem;
		}
		catch( ClientException $e ) {
			throw new serviceException( 'File not found', $e->getCode(), $e );
		}
		catch( GuzzleException $e ) {
			throw new serviceException( 'Error getting files from Microsoft', 500, $e );
		}
		catch( GraphException $e ) {
			throw new serviceException( $e->getMessage(), 500, $e );
		}
	}

	/**
	 * @param string $itemId Microsoft file id
	 * @param string $tmpPath Path to download the file into
	 *
	 * @return string File path name to tmp local copy of file
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function downloadDriveItemById( string $itemId, string $tmpPath='' ): string {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//get file list
		try {
			$graph = new \Microsoft\Graph\Graph();
			$graph->setAccessToken( $accessToken );

			/** @var \Microsoft\Graph\Model\DriveItem $driveItem */
			$driveItem = $graph->createRequest( "GET", "/drives/" . $this->config->driveId . '/items/' . $itemId )->setReturnType( \Microsoft\Graph\Model\DriveItem::class )->execute();

			//download the file, store it temporarily, serve it to user, delete file
			$path = rtrim( $tmpPath, '/' ) . '/' . $itemId;
			if(!file_exists($tmpPath)) {
				mkdir($tmpPath);
			}
			if(!file_exists($path)) {
				mkdir($path);
			}
			$filePathName = $path . '/' . $driveItem->getName();

			if(file_exists($filePathName)) {
				return $filePathName;
			}

			/** @var \Microsoft\Graph\Model\DriveItem $driveItem */
			$graph->createRequest( "GET", "/drives/" . $this->config->driveId . '/items/' . $itemId . '/content' )
			      ->download( $filePathName );

			return $filePathName;
		}
		catch( ClientException $e ) {
			throw new serviceException( 'File not found', $e->getCode(), $e );
		}
		catch( GuzzleException $e ) {
			throw new serviceException( 'Error getting files from Microsoft', 500, $e );
		}
		catch( GraphException $e ) {
			throw new serviceException( $e->getMessage(), 500, $e );
		}
	}



	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function moveItem( string $itemIdToMove, string $newParentDirItemId ) : \Microsoft\Graph\Model\DriveItem {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//get file list
		try {
			$graph = new \Microsoft\Graph\Graph();
			$graph->setAccessToken( $accessToken );

			/** @var \Microsoft\Graph\Model\DriveItem $driveItem */
			$driveItem = $graph->createRequest( "PATCH", "/drives/" . $this->config->driveId . "/items/" . $itemIdToMove )
				->attachBody( [ 'parentReference'=>[ 'id'=>$newParentDirItemId ]] )
				->setReturnType( \Microsoft\Graph\Model\DriveItem::class )->execute();

			return $driveItem;
		}
		catch( ClientException $e ) {
			throw new serviceException( $e->getMessage(), $e->getCode(), $e );
		}
		catch( GuzzleException $e ) {
			throw new serviceException( 'Error communicating with storage provider', 500, $e );
		}
		catch( GraphException $e ) {
			throw new serviceException( $e->getMessage(), 500, $e );
		}
	}


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function renameItem( string $itemIdToRename, string $newName ) : \Microsoft\Graph\Model\DriveItem {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//get file list
		try {
			$graph = new \Microsoft\Graph\Graph();
			$graph->setAccessToken( $accessToken );

			/** @var \Microsoft\Graph\Model\DriveItem $driveItem */
			$driveItem = $graph->createRequest( "PATCH", "/drives/" . $this->config->driveId . "/items/" . $itemIdToRename )
				->attachBody( [ 'name'=>$newName ] )
				->setReturnType( \Microsoft\Graph\Model\DriveItem::class )->execute();

			return $driveItem;
		}
		catch( ClientException $e ) {
			throw new serviceException( $e->getMessage(), $e->getCode(), $e );
		}
		catch( GuzzleException $e ) {
			throw new serviceException( 'Error communicating with storage provider', 500, $e );
		}
		catch( GraphException $e ) {
			throw new serviceException( $e->getMessage(), 500, $e );
		}
	}

	/**
	 * @param  string    $base64EncodedContent
	 * @param  string    $contentType
	 * @param  string    $fileName
	 * @param  string[]  $uploadPathParts  Ex: [ '2021-0001', 'Building 1', 'Inspections' ] will turn into {root}/2021-0001/Building 1/Inspections
	 * @param  string    $conflictBehavior The conflict resolution behavior for actions that create a new item. You can use the values fail, replace, or rename. The default for PUT is replace.
	 *
	 * @return \andrewsauder\microsoftServices\components\upload
	 */
	public function uploadBase64EncodedContent( string $base64EncodedContent, string $contentType, string $fileName, array $uploadPathParts=[], string $conflictBehavior='replace' ) : \andrewsauder\microsoftServices\components\upload {
		$tempFileName = tempnam(sys_get_temp_dir(), 'MicrosoftServicesFile');
		file_put_contents($tempFileName, base64_decode($base64EncodedContent));
		if($tempFileName===false) {
			$response = new \andrewsauder\microsoftServices\components\upload();
			$response->errors[] = 'Unable to write to tmp directory at '.sys_get_temp_dir();
			return $response;
		}

		return $this->upload( $tempFileName, $fileName, $uploadPathParts, $conflictBehavior );
	}

		/**
	 * @param  string    $serverFullFilePath
	 * @param  string    $fileName
	 * @param  string[]  $uploadPathParts  Ex: [ '2021-0001', 'Building 1', 'Inspections' ] will turn into {root}/2021-0001/Building 1/Inspections
	 * @param  string    $conflictBehavior The conflict resolution behavior for actions that create a new item. You can use the values fail, replace, or rename. The default for PUT is replace.
	 *
	 * @return \andrewsauder\microsoftServices\components\upload
	 */
	public function upload( string $serverFullFilePath, string $fileName, array $uploadPathParts=[], string $conflictBehavior='replace' ) : \andrewsauder\microsoftServices\components\upload {
		$response = new \andrewsauder\microsoftServices\components\upload();

		//get or user application access token
		try {
			$accessToken = $this->getMicrosoftAccessToken();
		}
		catch( serviceException $e ) {
			$response->errors[] = 'Invalid configuration: '.$e->getMessage();
			return $response;
		}

		//MICROSOFT UPLOAD
		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		$fileEndpoint = "/drives/" . $this->config->driveId . '/root:/' . $this->rootBasePath . implode( '/', $uploadPathParts ) . '/' . $fileName;

		$fileSize = filesize( $serverFullFilePath );

		try {
			//if less than 4 mb, simple upload
			if( $fileSize <= 4194304 ) {
				$driveItem = $graph->createRequest( "PUT", $fileEndpoint . ":/content?@microsoft.graph.conflictBehavior=".$conflictBehavior )->attachBody( file_get_contents( $serverFullFilePath ) )->setReturnType( \Microsoft\Graph\Model\DriveItem::class )->execute();
			}
			//larger than 4 mb, upload in chunks
			else {
				//1. create upload session
				$graphBody = [
					"@microsoft.graph.conflictBehavior" => $conflictBehavior,
					"description"                       => "",
					"fileSystemInfo"                    => [ "@odata.type" => "microsoft.graph.fileSystemInfo" ],
					"name"                              => $fileName,
				];

				$uploadSession = $graph->createRequest( "POST", $fileEndpoint . ":/createUploadSession" )->attachBody( $graphBody )->setReturnType( \Microsoft\Graph\Model\UploadSession::class )->execute();

				//2. upload bytes
				$fragSize       = 1024 * 1024 * 4;
				$graphUrl       = $uploadSession->getUploadUrl();
				$numFragments   = ceil( $fileSize / $fragSize );
				$bytesRemaining = $fileSize;
				$i              = 0;

				if( $stream = fopen( $serverFullFilePath, 'r' ) ) {
					while( $i < $numFragments ) {
						$chunkSize = $numBytes = $fragSize;
						$start     = $i * $fragSize;
						$end       = $i * $fragSize + $chunkSize - 1;
						$offset    = $i * $fragSize;
						if( $bytesRemaining < $chunkSize ) {
							$chunkSize = $numBytes = $bytesRemaining;
							$end       = $fileSize - 1;
						}

						// get contents using offset
						$data = stream_get_contents( $stream, $chunkSize, $offset );

						$content_range  = "bytes " . $start . "-" . $end . "/" . $fileSize;
						$headers        = [
							"Content-Length" => $numBytes,
							"Content-Range"  => $content_range
						];
						$uploadByte     = $graph->createRequest( "PUT", $graphUrl )->addHeaders( $headers )->attachBody( $data )->setReturnType( \Microsoft\Graph\Model\UploadSession::class )->setTimeout( "1000" )->execute();
						$bytesRemaining = $bytesRemaining - $chunkSize;
						$i++;
					}
					fclose( $stream );
				}

				$driveItem = $graph->createRequest( "GET", $fileEndpoint )->setReturnType( \Microsoft\Graph\Model\DriveItem::class )->execute();
			}

			$response->files[] = $driveItem;
		}
		catch( GraphException|GuzzleException $e ) {
			$response->errors[] = new \andrewsauder\microsoftServices\components\envelope( $e->getCode(), true, $fileName . ' did not upload. ' . $e->getMessage() );
		}

		return $response;
	}


	/**
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function delete( string $itemId ) {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//Microsoft Delete
		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		$itemEndpoint = "/drives/" . $this->config->driveId . "/items/" . $itemId;

		try {
			$deleteRequest = $graph->createRequest( "DELETE", $itemEndpoint )->execute();
		}
		catch( GuzzleException|GraphException $e ) {
			throw new serviceException( $e->getMessage(), 500, $e );
		}

		return $deleteRequest;
	}


	/**
	 * @param  string    $folderName
	 * @param  string[]  $basePathParts  Ex: [ '2021-0001', 'Building 1', 'Inspections' ] will turn into {root}/2021-0001/Building 1/Inspections
	 *
	 * @return \Microsoft\Graph\Model\DriveItem
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	public function createFolder( string $folderName, array $basePathParts=[] ) {
		//get application access token
		$accessToken = $this->getMicrosoftAccessToken();

		//Microsoft Delete
		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );


		$body = [
			'name'                              => $folderName,
			'folder'                            => (object) [],
			'@microsoft.graph.conflictBehavior' => 'fail'
		];

		try {
			/** @var \Microsoft\Graph\Model\DriveItem $driveItem */
			$driveItem = $graph->createRequest( "POST", "/drives/" . $this->config->driveId . '/root:/' . $this->rootBasePath . implode( '/', $basePathParts ) . ":/children" )->attachBody( $body )->setReturnType( \Microsoft\Graph\Model\DriveItem::class )
			                      ->execute();
		}
		catch( \Exception|\GuzzleHttp\Exception\GuzzleException $e ) {
			throw new serviceException( $this->rootBasePath . implode( '/', $basePathParts ) . '/' . $folderName . ' not created', 500, $e );
		}

		return $driveItem;
	}


	/**
	 * @param          $accessToken
	 * @param  string  $path
	 *
	 * @return \Microsoft\Graph\Model\DriveItem[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	private function getMicrosoftDriveItems( $accessToken, string $path, bool $recursive=true ) : array {
		//get file list
		try {
			$graph = new \Microsoft\Graph\Graph();
			$graph->setAccessToken( $accessToken );

			//get all the project folders
			/** @var \Microsoft\Graph\Model\DriveItem[] $driveItems */
			$driveItems = $graph->createRequest( "GET", '/drives/' . $this->config->driveId . '/root:/' . $this->rootBasePath . $path . ':/children' )->setReturnType( \Microsoft\Graph\Model\DriveItem::class )->execute();

			if($recursive) {
				foreach( $driveItems as $i => $driveItem ) {
					if( $driveItem->getFolder() !== null ) {
						$children = $this->getMicrosoftDriveItems( $accessToken, $path . '/' . $driveItem->getName() );
						$driveItems[ $i ]->setChildren( $children );
					}
				}
			}

			return $driveItems;
		}
		catch( ClientException $e ) {
			throw new serviceException( 'Folder not found', $e->getCode(), $e );
		}
		catch( GuzzleException $e ) {
			throw new serviceException( 'Error getting files from Microsoft', 500, $e );
		}
		catch( GraphException $e ) {
			throw new serviceException( $e->getMessage(), 500, $e );
		}
	}


	/**
	 * @param          $accessToken
	 * @param string   $id
	 * @param bool     $recursive
	 *
	 * @return \Microsoft\Graph\Model\DriveItem[]
	 * @throws \andrewsauder\microsoftServices\exceptions\serviceException
	 */
	private function getMicrosoftDriveItemsById( $accessToken, string $id, bool $recursive=true ) : array {
		//get file list
		try {
			$graph = new \Microsoft\Graph\Graph();
			$graph->setAccessToken( $accessToken );

			//get all the project folders
			/** @var \Microsoft\Graph\Model\DriveItem[] $driveItems */
			$driveItems = $graph->createRequest( "GET", '/drives/' . $this->config->driveId . '/items/' . $id . '/children' )->setReturnType( \Microsoft\Graph\Model\DriveItem::class )->execute();

			if($recursive) {
				foreach( $driveItems as $i => $driveItem ) {
					if( $driveItem->getFolder() !== null ) {
						$children = $this->getMicrosoftDriveItemsById( $accessToken, $driveItem->getId() );
						$driveItems[ $i ]->setChildren( $children );
					}
				}
			}

			return $driveItems;
		}
		catch( ClientException $e ) {
			throw new serviceException( 'Folder not found', $e->getCode(), $e );
		}
		catch( GuzzleException $e ) {
			throw new serviceException( 'Error getting files from Microsoft', 500, $e );
		}
		catch( GraphException $e ) {
			throw new serviceException( $e->getMessage(), 500, $e );
		}
	}


	/**
	 * @param            $accessToken
	 * @param  string[]  $basePath
	 *
	 * @return \Microsoft\Graph\Model\DriveItem[]
	 */
	private function createMicrosoftDirectories( $accessToken, array $basePath ) : array {
		$driveItems = [];

		$graph = new \Microsoft\Graph\Graph();
		$graph->setAccessToken( $accessToken );

		$path = '';
		foreach( $basePath as $i => $directoryName ) {
			$body = [
				'name'                              => $directoryName,
				'folder'                            => (object) [],
				'@microsoft.graph.conflictBehavior' => 'fail'
			];

			try {
				/** @var \Microsoft\Graph\Model\DriveItem $driveItem */
				$driveItems[] = $graph->createRequest( "POST", "/drives/" . $this->config->driveId . '/root:/' . $this->rootBasePath . $path . ":/children" )->attachBody( $body )->setReturnType( \Microsoft\Graph\Model\DriveItem::class )
				                      ->execute();
			}
			catch( \Exception|\GuzzleHttp\Exception\GuzzleException $e ) {
				error_log( $path . '/' . $directoryName . ' not created - already exists?' );
				error_log( $e );
			}

			//build out the path with each iteration since the array is the directory structure
			//ie if $basePath array = [ parent folder, child folder, grandchild folder ], $path becomes "/parent folder/child folder" after the second index
			$path .= '/' . $directoryName;
		}

		return $driveItems;
	}

}
