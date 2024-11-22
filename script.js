Office.initialize = function () {
	console.log("Office initialized");
};

function runJavaScript(event) {
	// Your JavaScript logic here
	console.log("Button clicked!");

	// Example: Display a simple message
	Office.context.mailbox.item.notificationMessages.addAsync("action", {
		type: "informationalMessage",
		message: "JavaScript executed successfully!",
		icon: "icon16",
		persistent: false,
	});

	// Call event.completed when done
	event.completed();
}
