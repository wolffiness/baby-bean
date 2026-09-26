function submitApplication() {
	const source = "submitApplication"
	try {
		const applicationSheet = getSheet(sheetNames.application)
		const name = fetchSheetData(applicationSheet, applicationRange.name)[0][0]
		const goalType = fetchSheetData(
			applicationSheet,
			applicationRange.metric,
		)[0][0]
		const goalAmount = fetchSheetData(
			applicationSheet,
			applicationRange.quantity,
		)[0][0]

		if (!validateForm(name, goalType, goalAmount)) {
			return
		}

		const userSheetName = `${name}`
		const response = checkExistingApplication(userSheetName, goalType)

		if (!response) {
			return
		}

		// Determine sheet settings according to goaltype
		let userSheetTemplate
		let ranges
		let formulas

		if (goalType == "Word count") {
			userSheetTemplate = getSheet(sheetNames.wordTemplate)
			ranges = templateRanges.word
			formulas = templateFormulas.word
		}

		if (goalType == "Task") {
			userSheetTemplate = getSheet(sheetNames.taskTemplate)
			ranges = templateRanges.task
			formulas = templateFormulas.task
		}

		// Create new sheet
		if (response !== "edit") {
			const totalSheets = spreadsheet.getNumSheets()
			spreadsheet.insertSheet(userSheetName, totalSheets, {
				template: userSheetTemplate,
			})
		}

		// customise sheet according to application
		const userSheet = getSheet(userSheetName)
		const nameRange = userSheet.getRange(ranges.name)
		nameRange.setValue(name)
		nameRange.protect().setWarningOnly(true)

		if (goalType == "Task")
			userSheet
				.getRange(ranges.todo)
				.setFormula(formulas.progress(ranges, goalAmount))

		userSheet.getRange(ranges.goal).setFormula(formulas.goal(ranges, goalAmount))
		userSheet
			.getRange(ranges.completion)
			.setFormula(formulas.completion(goalAmount))
		userSheet.getRange(ranges.daily).setFormula(formulas.daily(goalAmount))
		userSheet.getRange(ranges.weekly).setFormula(formulas.weekly(goalAmount))

		// remove application data
		applicationSheet.getRange(applicationRange.full).clearContent()

		createTrigger(statCalculations.participants.trigger_function, 5)

		logMessage(
			logTypes.info,
			`Finished ${response == "edit" ? "editing" : "creating"} sheet according to the application.`,
			source,
		)
		SpreadsheetApp.getUi().alert(
			`Finished ${response == "edit" ? "editing" : "creating"} sheet according to the application.`,
		)
	} catch (err) {
		logMessage(logTypes.err, `An error has occured: ${err.stack}`, source)
		SpreadsheetApp.getUi().alert(
			`An error has occured, please contact the script creator for help: ${err.stack}`,
		)
	}
}

function validateForm(name, goalType, goalAmount) {
	const source = "validateForm"

	if (!name || !goalType || !goalAmount) {
		logMessage(logTypes.warning, `Please fill in all application fields.`, source)
		logMessage(
			logTypes.debug,
			`name: ${name}. goalType: ${goalType}. goalAmount: ${goalAmount}`,
			source,
		)
		SpreadsheetApp.getUi().alert(`Please fill in all application fields.`)
		return false
	}

	if (
		(goalType !== "Word count" && goalType !== "Task") ||
		typeof goalAmount !== "number"
	) {
		logMessage(
			logTypes.warning,
			`Please make sure the goal amount is a number and the goal type is according to the dropdown.`,
			source,
		)
		logMessage(
			logTypes.debug,
			`goalType: ${goalType}. goalAmount: ${goalAmount}`,
			source,
		)
		SpreadsheetApp.getUi().alert(
			`Please make sure the goal amount is a number and the goal type is according to the dropdown.`,
		)
		return false
	}

	return true
}

function checkExistingApplication(userSheetName, goalType) {
	const source = "checkExistingApplication"

	if (spreadsheet.getSheetByName(userSheetName) !== null) {
		const response = SpreadsheetApp.getUi().alert(
			"Application already exists",
			"This name has already applied. Do you want to edit your existing application instead?",
			SpreadsheetApp.getUi().ButtonSet.YES_NO,
		)

		if (response == "NO") {
			logMessage(
				logTypes.warning,
				`This name has already applied, please choose a different one.`,
				source,
			)
			return false
		}

		logMessage(
			logTypes.info,
			`Choice has been made to edit existing application.`,
			source,
		)

		// check goalType
		const sheetTitle = getSheet(userSheetName).getRange("B2").getValue()
		const oldGoalType =
			sheetTitle == "Your accomplishments" ? "Task" : "Word count"
		logMessage(
			logTypes.debug,
			`sheetTitle: ${sheetTitle}. oldGoalType: ${oldGoalType}`,
			source,
		)

		if (oldGoalType !== goalType) {
			logMessage(
				logTypes.warning,
				`You have applied with a different goal type than your existing application. If you wish to change your goal type you will have to reapply instead.`,
				source,
			)
			SpreadsheetApp.getUi().alert(
				`You have applied with a different goal type than your existing application. If you wish to change your goal type you will have to reapply instead.`,
			)
		}

		return "edit"
	}

	return true
}
