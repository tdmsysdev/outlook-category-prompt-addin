function onMessageSendHandler(event) {
  event.completed({
    allowEvent: false,
    errorMessage: "Please assign at least one category (e.g., Project or Client) before sending.\n\nUse the Categorize button in the message ribbon.",
    sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser
  });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);