const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()

const sheetNames = {
	application: "Apply for Baby Bean",
	dashboard: "Dashboard",
	logs: "Logs",
	devLogs: "Dev Logs",
	template: "Template",
	wordTemplate: "Word Count Template",
	taskTemplate: "Task Template",
}

const statCalculations = {
	participants: {
		trigger_function: "participantsUpdateStat",
		cell: "C3",
		calculation: '=countStat("sheet")',
	},
	completion: {
		trigger_function: "completionUpdateStat",
		cell: "C4",
		calculation: '=CONCAT(ROUND(averageStat("completion")), "%")',
	},
}

const templateRanges = {
	name: "C3",
	goal: "D20",
	completion: "D21",
	daily: "D22",
	weekly: "D23",
}

const templateFormulas = {
	goal: (ranges, goalAmount) =>
		`=CONCAT(TEXT(SUMIFS(Logs!G:G, Logs!A:A, ${ranges.name}), "#,##0"), " / " & TEXT(${goalAmount}, "#,##0"))`,
	completion: (goalAmount) =>
		`=CONCAT(ROUND(((SUBSTITUTE(D19, " / " & TEXT(${goalAmount}, "#,##0"), "")) / ${goalAmount}) * 100, 0), "%")`,
	daily: (goalAmount) =>
		`=TEXT(MAX(0, ROUND((${goalAmount} - SUBSTITUTE(D19, " / " & TEXT(${goalAmount}, "#,##0"), "")) / (EOMONTH(TODAY(),0) - TODAY() + 1), 0)), "#,##0")`,
	weekly: (goalAmount) =>
		`=TEXT(MAX(0, ROUND((${goalAmount} - SUBSTITUTE(D19, " / " & TEXT(${goalAmount}, "#,##0"), "")) / MAX(1, ROUNDUP(((EOMONTH(TODAY(),0) - TODAY() + 1) / 7), 0)), 0)), "#,##0")`,
}

const applicationRange = {
	full: "B6:B8",
	name: "B6",
	metric: "B7",
	quantity: "B8",
}

const logSessionRange = {
	name: "C3",
	date: "E3",
	timeStart: "C4",
	timeEnd: "E4",
	countStart: "C5",
	countEnd: "E5",
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
		// .addItem('Update stats', 'initUpdateStats')
		.addItem("Update stats", "updateAllStats")
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

function updateStat(key = "participants") {
	const dashboardSheet = getSheet(sheetNames.dashboard)
	dashboardSheet.getRange(statCalculations[key].cell).clearContent()
	dashboardSheet
		.getRange(statCalculations[key].cell)
		.setFormula(statCalculations[key].calculation)

	SpreadsheetApp.flush()
	logMessage(logTypes.info, `Updated ${key} cell.`, "updateStat")
}

function updateAllStats() {
	Object.keys(statCalculations).forEach((key) => {
		updateStat(key)
	})
}

function participantsUpdateStat() {
	deleteExistingTriggers(statCalculations.participants.trigger_function)
	updateStat("participants")
}

function completionUpdateStat() {
	deleteExistingTriggers(statCalculations.completion.trigger_function)
	updateStat("completion")
}

function tasksCompletedUpdateStat() {
	deleteExistingTriggers(statCalculations.tasks_completed.trigger_function)
	updateStat("tasks_completed")
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
