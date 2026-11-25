# Microsoft Services

PHP wrapper for easy implementation of Microsoft Graph services

## Requirement
| Version | PHP | Microsoft Graph SDK |
| --- | --- | --- |
| >=2.0 | >=8.3 | 3.x |
| >=1.6 | >=8.3 | 1.x |
| 1.4.3 to 1.5.x | >=8.2 | 1.x |
| 1.2 to 1.4.2 | >=8.1 | 1.x |

To upgrade from 1.x, see [UPGRADING.md](UPGRADING.md).

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

## Authentication

Each service takes the config and an optional user access token. The token decides how the service signs in
to Microsoft Graph:

| User access token | `onBehalfOfFlow` | Graph token |
| --- | --- | --- |
| None | Any | An application token (client credentials) |
| Provided | `true` (default) | A token from the on-behalf-of exchange of the user token |
| Provided | `false` | The user token as it is |

The service requests its token once, at its first Graph call, and reuses it for every later call. An
application token and an on-behalf-of token use the `https://graph.microsoft.com/.default` scope.

The methods return the models of the Microsoft Graph SDK, for example
`\Microsoft\Graph\Generated\Models\DriveItem`. A failed Graph call throws a
`\andrewsauder\microsoftServices\exceptions\serviceException`. Its code is the HTTP status of the Graph
response, or 500 when there is no response.

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
$microsoftFiles->delete( $itemId );
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

try {
    $microsoftMail->send( 'to@example.com', 'Subject', 'HTML compatible message', 'from@example.com' );
}
catch( \andrewsauder\microsoftServices\exceptions\serviceException $e ) {
    error_log( 'Failed: ' . $e->getMessage() );
}
```

`send()` puts every attachment in one `sendMail` request, and Graph rejects a request over about 4 MB.
For larger attachments, use a draft.


### Send Email With Large Attachments

`createDraft()` creates the draft, then attaches each file in its own request. A file up to 3 MB goes in
one `POST`. A larger file, up to 150 MB, goes through an upload session in ranges of about 3 MB. The
total is limited only by the mailbox's message size limit: 35 MB by default in Exchange Online.

If an attachment fails, `createDraft()` deletes the draft and throws a `serviceException`.

```php
$microsoftMail = new \andrewsauder\microsoftServices\mail( $config );
$microsoftMail->addAttachment( '/tmp/report.pdf' );
$microsoftMail->addAttachment( '/tmp/chain-of-custody.pdf' );

$draft = $microsoftMail->createDraft( [ 'to@example.com' ], 'Subject', 'HTML compatible message', 'from@example.com', [ 'cc@example.com' ] );
$microsoftMail->sendDraft( $draft->getId(), 'from@example.com' );
```

The application needs the `Mail.ReadWrite` permission to create the draft and attach files, and
`Mail.Send` to send it.


### Reply Into a Thread

Continue an existing conversation instead of starting a new one. `replyIntoThread()` creates a reply
draft anchored on a message you previously sent (so it inherits that conversation and its `RE:` subject),
sets the recipients and body you provide, and sends it. Store the returned message's immutable id as the
anchor for the next reply, and its conversation id for inbound matching.

Provide the `from` mailbox, the immutable id of the last message you sent in the thread, and the
recipients/body. Do not set a subject - the reply inherits it.

```php
$microsoftMail = new \andrewsauder\microsoftServices\mail( $config );

$sent = $microsoftMail->replyIntoThread(
    'helpdesk@example.com',   // from (owns the anchor message)
    $lastMessageId,           // immutable id of the last message sent in the thread
    [ 'to@example.com' ],     // to
    [],                       // bcc
    'HTML compatible message',
    [ 'cc@example.com' ]      // cc (optional)
);

$nextAnchor     = $sent->getId();              // store as the next thread anchor
$conversationId = $sent->getConversationId();  // store for inbound reply matching
```

If the anchor message no longer exists (deleted), `replyIntoThread()` throws a `serviceException` - catch
it and fall back to `createDraft()` / `sendDraft()` to start a fresh thread.

Only the immutable id is needed to reply. If you have lost it but still know the conversation id, recover
an anchor with:

```php
$latest = $microsoftMail->getLatestMessageInConversation( 'helpdesk@example.com', $conversationId );
$anchorId = $latest?->getId();
```

You can also create just the draft (without sending) with `createReplyDraft( $messageId, $from )`.


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

foreach( $attachments as $attachment ) {
    if( $attachment instanceof \Microsoft\Graph\Generated\Models\FileAttachment ) {
        //the stream holds the base64 text of the file
        file_put_contents( '/tmp/' . $attachment->getName(), base64_decode( (string)$attachment->getContentBytes() ) );
    }
}
```


### Move Message
```php
$microsoftMail = new \andrewsauder\microsoftServices\mail( $config );
$movedMessage = $microsoftMail->moveMessage( 'joeschmoe@example.com', 'message-id', 'archive' );
```


### Delete Message
```php
$microsoftMail = new \andrewsauder\microsoftServices\mail( $config );
$microsoftMail->deleteMessage( 'joeschmoe@example.com', 'message-id' );
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


### Get User by Id
```php
$microsoftUserService = new \andrewsauder\microsoftServices\user( $config );
$users = $microsoftUserService->getUserById( '15bd6895-bf60-4125-a1d2-affb7e0de5d8' );
```


### Get Users By Advanced Filter
See https://learn.microsoft.com/en-us/graph/filter-query-parameter?tabs=http for filter details
```php
$microsoftUserService = new \andrewsauder\microsoftServices\user( $config );
$users = $microsoftUserService->getUsersByFilter( 'startswith(userPrincipalName,"andrew")' );
```
