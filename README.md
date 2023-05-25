# Microsoft Services

PHP wrapper for easy implementation of Microsoft Graph services

## Requirement
Version >=1.2 requires PHP >=8.1

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
If no user token is provided, the application token will be used.

If the application token is being used, verify that the Azure application has correct Mail.X permissions for the email 
address being used. To limit application access to only certain mailboxes, use ExchangeOnline Powershell to apply access
policy. More info https://learn.microsoft.com/en-us/powershell/module/exchange/new-applicationaccesspolicy?view=exchange-ps

If a user access token is provided when creating the service (`mail($config, 'user-access-token-string')`), verify that 
the user has 'send on behalf' of or 'send as' permissions configured properly in Office 365. 


### Send Email

If the from address is not provided, the default from address in the config will be used.

```php
$microsoftMail = new \andrewsauder\microsoftServices\mail( $config );
$microsoftMail->addAttachment( 'C:\tmp\testFile.txt' );
$rsp = $microsoftMail->send( 'to@example.com', 'Subject', 'HTML compatible message', 'from@example.com' );

if( $rsp->getStatus() < 200 || $rsp->getStatus() >= 300 ) {
    error_log( 'Failed' );
}
```


### Get All Messages
```php
$microsoftMail = new \andrewsauder\microsoftServices\mail( $config );
$messages = $microsoftMail->getAllMessages( 'joeschmoe@example.com' );
```

### Get All Messages from Specific Folder
```php
$microsoftMail = new \andrewsauder\microsoftServices\mail( $config );
$messages = $microsoftMail->getMessagesInFolder( 'joeschmoe@example.com', 'mail-folder-id' );
```


### Get All Folders
```php
$microsoftMail = new \andrewsauder\microsoftServices\mail( $config );
$folders = $microsoftMail->getFolders( 'joeschmoe@example.com' );
```


### Get Attachments for Message
```php
$microsoftMail = new \andrewsauder\microsoftServices\mail( $config );
$attachments = $microsoftMail->getAttachments( 'joeschmoe@example.com', 'message-id' );
```


### Delete Message
```php
$microsoftMail = new \andrewsauder\microsoftServices\mail( $config );
$graphResponse = $microsoftMail->deleteMessage( 'joeschmoe@example.com', 'message-id' );
```

## User Usage

### Get All Users in Organization

```php
$microsoftUserService = new \andrewsauder\microsoftServices\user( $config );
$users = $microsoftUserService->allUsersInOrganization();
```

### Get User by User Principal Name
```php
$microsoftUserService = new \andrewsauder\microsoftServices\user( $config );
$users = $microsoftUserService->getUserByUserPrincipalName( 'andrew@sauder.software' );
```


### Get User by External Id
```php
$microsoftUserService = new \andrewsauder\microsoftServices\user( $config );
$users = $microsoftUserService->getUserByExternalId( '15bd6895-bf60-4125-a1d2-affb7e0de5d8' );
```


### Get Users By Advanced Filter
```php
$microsoftUserService = new \andrewsauder\microsoftServices\user( $config );
$users = $microsoftUserService->getUsersByFilter( 'startswith(userPrincipalName,"andrew")' );
```
