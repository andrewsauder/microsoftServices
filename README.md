# Microsoft Services

PHP wrapper for easy implementation of Microsoft Graph services

## Installation

`composer require andrewsauder/microsoft-services`

## Service Configuration

```php
$config = new \andrewsauder\microsoftServices\config();
$config->clientId = '{Azure Application ID}';
$config->clientSecret = '{Azure Client Secret}';  //certificates not yet supported
$config->tenant = 'example.com';
$config->driveId = '';                            //required if using the files service - cay be found using Graph explorer
$config->fromAddress = 'noreply@example.com';     //required if using mail service - this is just a default
```

## Files Usage

### Get List of Files

#### From Root Directory

```php
$microsoftFiles = new \andrewsauder\microsoftServices\files( $config );
$files = $microsoftFiles->list();
```

#### From Subdirectory

```php
//example subdirectory: {root}/2021-0001/Building 1/Inspections
$microsoftFiles = new \andrewsauder\microsoftServices\files( $config );
$files = $microsoftFiles->list( [ '2021-0001', 'Building 1', 'Inspections' ] );
```

### Upload File

#### Into Root Directory

```php
$microsoftFiles = new \andrewsauder\microsoftServices\files( $config );
$uploadFileResponse = $microsoftFiles->upload( 'C:\tmp\testFile.txt', 'testFile.txt' );
```

#### Into Subdirectory

```php
//example subdirectory: {root}/2021-0001/Building 1/Inspections
$microsoftFiles = new \andrewsauder\microsoftServices\files( $config );
$uploadFileResponse = $microsoftFiles->upload( 'C:\tmp\testFile.txt', 'testFile.txt', [ '2021-0001', 'Building 1', 'Inspections' ] );
```

### Delete File

```php
$microsoftFiles = new \andrewsauder\microsoftServices\files( $config );
$deleteResponse = $microsoftFiles->delete( $itemId );
```

## Mail Usage

### Send Email

If the from address is not provided, the default from address in the config will be used. If sending from a different
account than the provided user token, make sure send on behalf of or send as permissions are defined in Office 365.

```php
$microsoftMail = new \gcgov\framework\services\microsoft\mail( $config );
$microsoftMail->addAttachment( 'C:\tmp\testFile.txt' );
$rsp = $microsoftMail->send( 'to@example.com', 'Subject', 'HTML compatible message', 'from@example.com' );

if( $rsp->getStatus() < 200 || $rsp->getStatus() >= 300 ) {
    error_log( 'Failed' );
}
```
