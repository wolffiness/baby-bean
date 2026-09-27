const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()

const sheetNames = {
	application: "Apply for Baby Bean",
	dashboard: "Dashboard",
	logs: "Logs",
	template: "Template",
	devLogs: "Dev Logs",
}

const goalTypes = ["Words", "Scenes", "Chapters", "Pages", "Lines"]

const templateRanges = {
	name: "C3",
	metric: "E3",
	goalAmount: "E4",
}

const templateFormulas = {
	goal: (ranges, goalAmount) =>
		`=CONCAT(TEXT(SUMIF(Logs!A:A, C3, Logs!G:G), "#,##0"), " / " & TEXT(${goalAmount}, "#,##0"))`,
	completion: (ranges, goalAmount) =>
		`=CONCAT(ROUND(((SUBSTITUTE(${ranges.goal}, " / " & TEXT(${goalAmount}, "#,##0"), "")) / ${goalAmount}) * 100, 0), "%")`,
	daily: (ranges, goalAmount) =>
		`=TEXT(MAX(0, ROUND((${goalAmount} - SUBSTITUTE(${ranges.goal}, " / " & TEXT(${goalAmount}, "#,##0"), "")) / (EOMONTH(TODAY(),0) - TODAY() + 1), 0)), "#,##0")`,
	weekly: (ranges, goalAmount) =>
		`=TEXT(MAX(0, ROUND((${goalAmount} - SUBSTITUTE(${ranges.goal}, " / " & TEXT(${goalAmount}, "#,##0"), "")) / MAX(1, ROUNDUP(((EOMONTH(TODAY(),0) - TODAY() + 1) / 7), 0)), 0)), "#,##0")`,
}

const applicationRange = {
	full: "B6:B8",
	name: "B6",
	metric: "B7",
	quantity: "B8",
}

const logSessionRange = {
	name: "C3",
	metric: "E3",
	date: "C4",
	timeStart: "C5",
	timeEnd: "E5",
	countStart: "C6",
	countEnd: "E6",
}

const logTypes = {
	debug: "DEBUG",
	info: "INFO",
	warning: "WARNING",
	error: "ERROR",
}

const colors = {
	INFO: "#d0f0c0",
	DEBUG: "#b1d8fb",
	WARNING: "#fff3cd",
	ERROR: "#f89694",
	DEFAULT: "#fcfefc",
}

function onOpen() {
	const ui = SpreadsheetApp.getUi()
	// DEVELOPER MENU
	ui
		.createMenu("Dev tools")
		.addItem("Clear logs", "initCleanLogs")
		.addToUi()
}

function getSheet(sheetName) {
	const source = "getSheet"
	const sheet = spreadsheet.getSheetByName(sheetName)
	if (!sheet) {
		logMessage(logTypes.error, `Sheet '${sheetName}' not found`, source)
	}

	return sheet
}

function fetchSheetData(sheet, range = null) {
	const source = "fetchSheetData"
	let data = []

	if (range !== null) {
		data = sheet.getRange(range).getValues()
	} else {
		data = sheet.getDataRange().getValues()
	}

	if (data.length === 0) {
		logMessage(
			logTypes.error,
			`Sheet '${sheet.getName()}' does not contain data.`,
			source,
		)
	}

	return data
}

function createTrigger(name, duration) {
	ScriptApp.newTrigger(name)
		.timeBased()
		.after(duration * 60000)
		.create()
	logMessage(
		logTypes.info,
		`Trigger created. Function ${name} will start after ${duration} minute.`,
		"createTrigger",
	)
}

function deleteExistingTriggers(triggerName) {
	const triggers = ScriptApp.getProjectTriggers()

	triggers.forEach((trigger) => {
		if (trigger.getHandlerFunction() === triggerName) {
			ScriptApp.deleteTrigger(trigger)
		}
	})
}

function backupSheet() {
	const source = "backupSheet"

	try {
		const originalFile = DriveApp.getFileById(spreadsheet.getId())
		const timestamp = Utilities.formatDate(
			new Date(),
			Session.getScriptTimeZone(),
			"yyyy-MM-dd HH:mm",
		)
		const backupName = `${spreadsheet.getName()} - Backup ${timestamp}`

		// Keep backups next to the original file rather than in Drive root
		const parentFolders = originalFile.getParents()
		const parentFolder = parentFolders.hasNext()
			? parentFolders.next()
			: DriveApp.getRootFolder()

		const backupFolderName = "Backups"
		const existingFolders = parentFolder.getFoldersByName(backupFolderName)
		const backupFolder = existingFolders.hasNext()
			? existingFolders.next()
			: parentFolder.createFolder(backupFolderName)

		const backupFile = originalFile.makeCopy(backupName, backupFolder)

		logMessage(
			logTypes.info,
			`Backup created: "${backupFile.getName()}".`,
			source,
		)
	} catch (err) {
		logMessage(logTypes.error, `Error creating backup: ${err.stack}`, source)
	}
}