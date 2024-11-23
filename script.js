Office.initialize = function () {
	console.log("Office initialized");
};

// function runJavaScript(event) {
// 	// Your JavaScript logic here
// 	console.log("Button clicked!");

// 	// Example: Display a simple message
// 	Office.context.mailbox.item.notificationMessages.addAsync("action", {
// 		type: "informationalMessage",
// 		message: "JavaScript executed successfully!",
// 		icon: "icon16",
// 		persistent: false,
// 	});

// 	// Call event.completed when done
// 	event.completed();
// }

async function runJavaScript(event) {
	try {
		await moveEmailToConversationFolder();
	} catch (error) {
		console.error(error);
		Office.context.mailbox.item.notificationMessages.addAsync("error", {
			type: "errorMessage",
			message: "Failed to move the email: " + error.message,
		});
	} finally {
		// Notify Outlook that the operation is complete
		event.completed();
	}
}

async function moveEmailToConversationFolder() {
	const mailbox = Office.context.mailbox;
	const item = mailbox.item;

	// Ensure the item is a message
	if (item.itemType !== Office.MailboxEnums.ItemType.Message) {
		throw new Error("This action can only be performed on email messages.");
	}

	const conversationId = item.conversationId;

	if (!conversationId) {
		throw new Error("No conversation ID found for this message.");
	}

	// Find all messages in the conversation
	const searchQuery = `conversationId:${conversationId}`;
	const options = {
		top: 10, // Adjust based on how many messages to retrieve
	};

	const messages = await searchMailbox(searchQuery, options);

	if (!messages || messages.length === 0) {
		throw new Error("No related messages found in the conversation.");
	}

	// Get the folderId of the first message in the conversation
	const folderId = messages[0].parentFolderId;

	if (!folderId) {
		throw new Error("Could not determine the folder of the conversation.");
	}

	// Move the current message
	await moveItem(item.itemId, folderId);

	// Display a success message
	Office.context.mailbox.item.notificationMessages.addAsync("success", {
		type: "informationalMessage",
		message: "Email moved to the conversation folder successfully.",
		icon: "icon16",
		persistent: false,
	});
}

function searchMailbox(query, options) {
	return new Promise((resolve, reject) => {
		const searchRequest = `
      <m:FindItem xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
        <m:ItemShape>
          <t:BaseShape>IdOnly</t:BaseShape>
        </m:ItemShape>
        <m:IndexedPageItemView MaxEntriesReturned="${options.top}" Offset="0" BasePoint="Beginning" />
        <m:Restriction>
          <t:Contains ContainmentMode="Substring" ContainmentComparison="IgnoreCase">
            <t:FieldURI FieldURI="item:ConversationId"/>
            <t:Constant Value="${query}" />
          </t:Contains>
        </m:Restriction>
      </m:FindItem>
    `;

		Office.context.mailbox.makeEwsRequestAsync(searchRequest, (result) => {
			if (result.status === Office.AsyncResultStatus.Succeeded) {
				const response = parseXml(result.value);
				const items =
					response.Body.ResponseMessages.FindItemResponseMessage.RootFolder
						.Items.Message;
				resolve(items || []);
			} else {
				reject(new Error(result.error.message));
			}
		});
	});
}

function moveItem(itemId, folderId) {
	return new Promise((resolve, reject) => {
		const moveRequest = `
      <m:MoveItem xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
        <m:ToFolderId>
          <t:FolderId Id="${folderId}" />
        </m:ToFolderId>
        <m:ItemIds>
          <t:ItemId Id="${itemId}" />
        </m:ItemIds>
      </m:MoveItem>
    `;

		Office.context.mailbox.makeEwsRequestAsync(moveRequest, (result) => {
			if (result.status === Office.AsyncResultStatus.Succeeded) {
				resolve();
			} else {
				reject(new Error(result.error.message));
			}
		});
	});
}

function parseXml(xmlString) {
	const parser = new DOMParser();
	return parser.parseFromString(xmlString, "text/xml");
}
