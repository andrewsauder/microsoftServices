# Upgrade guide

## From 1.x to 2.0

Version 2.0 moves the library from the Microsoft Graph SDK 1.x (`microsoft/microsoft-graph ^1.110`) to
the SDK 3.x (`^3.11`). The SDK 3.x is a new, generated SDK. Most changes in this guide come from it.

The PHP requirement does not change: PHP 8.3 or later.

### Update the model class names

The methods return the models of the SDK 3.x. Each model moves to a new namespace.

| 1.x | 2.0 |
| --- | --- |
| `\Microsoft\Graph\Model\User` | `\Microsoft\Graph\Generated\Models\User` |
| `\Microsoft\Graph\Model\DriveItem` | `\Microsoft\Graph\Generated\Models\DriveItem` |
| `\Microsoft\Graph\Model\Message` | `\Microsoft\Graph\Generated\Models\Message` |

The same rule applies to every other model, for example `Calendar`, `Event`, `MailFolder`, `Attachment`,
`ItemPreviewInfo`, and `Thumbnail`.

Most getters keep their names, for example `getId()`, `getName()`, and `getSubject()`. These getters change:

- An enum property returns an enum object, not a string. For example,
  `$message->getBody()->getContentType()` returns a `\Microsoft\Graph\Generated\Models\BodyType`. Call
  `->value()` on it to get the string.
- `FileAttachment::getContentBytes()` returns a `\Psr\Http\Message\StreamInterface`. The stream holds the
  base64 text of the file. Cast it to a string, then call `base64_decode()`.

### Update the methods that returned a `GraphResponse`

The SDK 3.x has no `GraphResponse`. These methods now return `void` and throw a `serviceException` when
Graph does not accept the request:

- `mail::send()`
- `mail::sendDraft()`
- `mail::deleteMessage()`
- `files::delete()`

`mail::moveMessage()` returns the moved `Message`.

Replace a status check:

```php
// 1.x
$rsp = $microsoftMail->send( 'to@example.com', 'Subject', 'Body' );
if( $rsp->getStatus() >= 300 ) {
    error_log( 'Failed' );
}

// 2.0
try {
    $microsoftMail->send( 'to@example.com', 'Subject', 'Body' );
}
catch( \andrewsauder\microsoftServices\exceptions\serviceException $e ) {
    error_log( 'Failed: ' . $e->getMessage() );
}
```

### Replace the removed methods

| Removed | Use instead |
| --- | --- |
| `user::getUserByEmail()` | `user::getUserByUserPrincipalName()`. The 1.x method always looked up the user principal name. |
| `files::getFile()` | `files::getDriveItem()` |
| `auth::getApplicationAccessToken()` | Nothing. The services get the application token themselves. |
| The `$nextLink` parameter of `user::allUsersInOrganization()` | Nothing. The method follows every next link itself. |

### Check the error codes

A `serviceException` now carries the HTTP status of the Graph response as its code, for example 404 or
409. In 1.x, many methods used 500 for every Graph error. The code is 500 only when there is no response.

### Check the access token scope

An application token and an on-behalf-of token now use the `https://graph.microsoft.com/.default` scope.
The token carries the Graph permissions that the app registration has consent for.

- In 1.x, the on-behalf-of exchange for Graph calls asked for `config->scope`. `config->scope` now applies
  to `auth::getAccessToken()` and `auth::verify()` only.
- In 1.x, the application token came from the Azure AD v1 token endpoint. It now comes from the Microsoft
  identity platform v2 endpoint.

A service now requests its token once and reuses it. In 1.x, each method call requested a new token.

### Check the changed results

These methods return different results in 2.0:

- `files::list()` and `files::listById()` return every child of each folder. In 1.x, they returned only
  the first page of children, at most 200 items.
- `user::getUsersByFilter()` returns every matching user. In 1.x, it returned only the first page.
- `calendar::getNCalendarEventsOccurringDuringOrAfter()` returns at most the number of events you ask
  for. In 1.x, it followed every next link, so it returned all the events after the date.
- `files::getDriveItemPreviewById()` sends the page number as `page`, the name that Graph reads. In 1.x,
  it sent `pageNumber`, and Graph ignored it.
- `files::getDriveItemThumbnailById()` accepts the sizes `small`, `medium`, `large`, and `source`. Any
  other size throws a `serviceException` with code 400.
- `files::uploadBase64EncodedContent()` deletes its temporary file after the upload.
