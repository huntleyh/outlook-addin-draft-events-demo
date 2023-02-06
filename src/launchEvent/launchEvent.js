/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

function onMessageAttachmentsChangedHandler(event) {
  console.log("Message attachments changed...");
  checkAttachmentsExist(event);
}

function onMessageRecipientsChangedHandler(event) {
  console.log("Message recipients changed...");
  checkAttachmentsExist(event);
}

function onMessageComposeHandler(event) {
  console.log("Composing new message or editing draft...");
  checkAttachmentsExist(event);
}

function checkAttachmentsExist(event) {
  console.log("Checking attatchments with external recipients.");
  attachmentExternalUsersValidation(event);
}

function attachmentExternalUsersValidation(event) {
  let externalUsers = [];
  let domainName = "@contoso.com";
  var panelId = "externalrecipients";
  Office.context.mailbox.item.getAttachmentsAsync(function (itemAttachments) {
    if (itemAttachments.value != null && itemAttachments.value.length > 0) {
      console.log(itemAttachments);
      Office.context.mailbox.item.to.getAsync(function (toAsyncResult) {
        if (toAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
          //to external users
          const msgTo = toAsyncResult.value;
          for (let i = 0; i < msgTo.length; i++) {
            console.log(msgTo[i].displayName + " (" + msgTo[i].emailAddress + ")");
            if (
              msgTo[i].emailAddress != null &&
              msgTo[i].emailAddress != undefined &&
              msgTo[i].emailAddress.trim() != "" &&
              //&& msgTo[i].recipientType.trim().toLowerCase() == "externaluser")
              msgTo[i].emailAddress.trim().toLowerCase().indexOf(domainName) <= -1
            )
              externalUsers.push("to:" + msgTo[i].emailAddress);
          }
          Office.context.mailbox.item.cc.getAsync(function (ccAsyncResult) {
            if (ccAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
              //cc external users
              const msgCC = ccAsyncResult.value;
              for (let i = 0; i < msgCC.length; i++) {
                console.log(msgCC[i].displayName + " (" + msgCC[i].emailAddress + ")");
                if (
                  msgCC[i].emailAddress != null &&
                  msgCC[i].emailAddress != undefined &&
                  msgCC[i].emailAddress.trim() != "" &&
                  //        && msgCC[i].recipientType.trim().toLowerCase() == "externaluser")
                  msgCC[i].emailAddress.trim().toLowerCase().indexOf(domainName) <= -1
                )
                  externalUsers.push("cc:" + msgCC[i].emailAddress);
              }
              Office.context.mailbox.item.bcc.getAsync(function (bccAsyncResult) {
                if (bccAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  //bcc external users
                  const msgBcc = bccAsyncResult.value;
                  for (let i = 0; i < msgBcc.length; i++) {
                    console.log(msgBcc[i].displayName + " (" + msgBcc[i].emailAddress + ")");
                    if (
                      msgBcc[i].emailAddress != null &&
                      msgBcc[i].emailAddress != undefined &&
                      msgBcc[i].emailAddress.trim() != "" &&
                      //&& msgBcc[i].recipientType.trim().toLowerCase() == "externaluser")
                      msgBcc[i].emailAddress.trim().toLowerCase().indexOf(domainName) <= -1
                    )
                      externalUsers.push("bcc:" + msgBcc[i].emailAddress);
                  }
                  let externalDomainNames = [];
                  for (let domainIndex = 0; domainIndex < externalUsers.length; domainIndex++) {
                    let externalUserEmail = externalUsers[domainIndex]
                      .replace("to:", "")
                      .replace("cc:", "")
                      .replace("bcc:", "");
                    var domIndex = externalUserEmail.indexOf("@");
                    var externalDomainName = externalUserEmail.substr(domIndex + 1);
                    if (externalDomainNames.indexOf(externalDomainName) === -1) {
                      externalDomainNames.push(externalDomainName);
                    }
                  }
                  if (externalUsers != null && externalUsers.length > 1 && externalDomainNames.length > 1) {
                    console.log(
                      "[WAIT!] Check that your recipients are correct. This email is going to external recipients with different email domains."
                    );
                    let externalUsersArray = [];
                    for (let userCount = 0; userCount < externalUsers.length; userCount++) {
                      let userEmail = externalUsers[userCount]
                        .replace("to:", "")
                        .replace("cc:", "")
                        .replace("bcc:", "");
                      let userThere = false;
                      for (
                        let existingUserCount = 0;
                        existingUserCount < externalUsersArray.length;
                        existingUserCount++
                      ) {
                        if (
                          externalUsersArray[existingUserCount].trim().toLowerCase() == userEmail.trim().toLowerCase()
                        ) {
                          userThere = true;
                          break;
                        }
                      }
                      if (!userThere) externalUsersArray.push(userEmail);
                    }
                    let externalUsersEmail = externalUsersArray.join(",");
                    let information =
                      "[WAIT!] Check that your recipients are correct. This email is going to external recipients with different email domains.";
                    let informationTruncated =
                      information.length > 150 ? information.slice(0, 150) + ".." : information;
                    console.log("Host name : " + Office.context.mailbox.diagnostics.hostName);
                    var paneldetails = {
                      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                      message: informationTruncated,
                      icon: "ErrorIcon",
                      persistent: false,
                    };
                    if (Office.context.mailbox.diagnostics.hostName == "OutlookWebApp") {
                      paneldetails = {
                        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
                        message: informationTruncated,
                      };
                    }
                    try {
                      Office.context.mailbox.item.notificationMessages.removeAsync(panelId, function (removeResult) {
                        console.log(removeResult);
                        Office.context.mailbox.item.notificationMessages.addAsync(
                          panelId,
                          paneldetails,
                          function (addResult) {
                            console.log(addResult);
                            event.completed({
                              allowEvent: true,
                            });
                            return;
                            //updateBody(event, information);
                          }
                        );
                      });
                    } catch (err) {
                      console.log(err);
                      event.completed({
                        allowEvent: true,
                      });
                      return;
                    }
                  } else {
                    Office.context.mailbox.item.notificationMessages.removeAsync(panelId, function (removeResult) {
                      console.log(removeResult);
                      event.completed({
                        allowEvent: true,
                      });
                      return;
                    });
                  }
                } else {
                  if (bccAsyncResult.error != null) {
                    console.error(bccAsyncResult.error);
                    Office.context.mailbox.item.notificationMessages.removeAsync(panelId, function (removeResult) {
                      console.log(removeResult);
                      event.completed({
                        allowEvent: true,
                      });
                      return;
                    });
                  } else {
                    event.completed({
                      allowEvent: true,
                    });
                    return;
                  }
                }
              });
            } else {
              if (ccAsyncResult.error != null) {
                console.error(ccAsyncResult.error);
                Office.context.mailbox.item.notificationMessages.removeAsync(panelId, function (removeResult) {
                  console.log(removeResult);
                  event.completed({
                    allowEvent: true,
                  });
                  return;
                });
              } else {
                event.completed({
                  allowEvent: true,
                });
                return;
              }
            }
          });
        } else {
          if (toAsyncResult.error != null) {
            console.error(toAsyncResult.error);
            Office.context.mailbox.item.notificationMessages.removeAsync(panelId, function (removeResult) {
              console.log(removeResult);
              event.completed({
                allowEvent: true,
              });
              return;
            });
          } else {
            event.completed({
              allowEvent: true,
            });
            return;
          }
        }
      });
    } else {
      Office.context.mailbox.item.notificationMessages.removeAsync(panelId, function (removeResult) {
        console.log(removeResult);
        event.completed({
          allowEvent: true,
        });
        return;
      });
    }
  });
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onMessageAttachmentsChangedHandler", onMessageAttachmentsChangedHandler);
  Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
  Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
}
