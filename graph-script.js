Office.initialize = function () {
	console.log("Office initialized");
};

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

	if (item.itemType !== Office.MailboxEnums.ItemType.Message) {
		throw new Error("This action can only be performed on email messages.");
	}

	const accessToken = await getAccessToken();
	const itemId = item.itemId;

	// Use Microsoft Graph to move the email
	const moveUrl = `https://graph.microsoft.com/v1.0/me/messages/${itemId}/move`;
	const targetFolderId = await getTargetFolderId(accessToken); // Logic to determine folder

	const response = await fetch(moveUrl, {
		method: "POST",
		headers: {
			Authorization: `Bearer ${accessToken}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ destinationId: targetFolderId }),
	});

	if (!response.ok) {
		throw new Error(`Failed to move email: ${response.statusText}`);
	}

	console.log("Email moved successfully!");
}

async function getAccessToken() {
	return new Promise((resolve, reject) => {
		Office.context.auth.getAccessTokenAsync(
			{ allowSignInPrompt: true },
			(result) => {
				if (result.status === Office.AsyncResultStatus.Succeeded) {
					resolve(result.value);
				} else {
					reject(new Error(result.error.message));
				}
			}
		);
	});
}

async function getTargetFolderId(accessToken) {
	// Logic to identify the target folder based on conversation ID or other criteria
	const folderName = "Target Folder Name"; // Replace with desired folder name
	const response = await fetch(
		"https://graph.microsoft.com/v1.0/me/mailFolders",
		{
			headers: {
				Authorization: `Bearer ${accessToken}`,
			},
		}
	);

	if (!response.ok) {
		throw new Error("Failed to fetch mail folders");
	}

	const data = await response.json();
	const folder = data.value.find((f) => f.displayName === folderName);

	if (!folder) {
		throw new Error("Target folder not found");
	}

	return folder.id;
}
