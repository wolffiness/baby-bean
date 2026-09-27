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

		// Create a new sheet from the single generic template
		if (response !== "edit") {
			const totalSheets = spreadsheet.getNumSheets()
			spreadsheet.insertSheet(userSheetName, totalSheets, {
				template: getSheet(sheetNames.template),
			})
		}

		const userSheet = getSheet(userSheetName)

		// Swap the [[METRIC]] placeholder for the goal type they picked
		userSheet
			.createTextFinder("[[METRIC]]")
			.matchEntireCell(false)
			.replaceAllWith(goalType)

		const nameRange = userSheet.getRange(templateRanges.name)
		nameRange.setValue(name)
		nameRange.protect().setWarningOnly(true)

		userSheet
			.getRange(templateRanges.sessions)
			.setFormula(templateFormulas.sessions(templateRanges))
		userSheet
			.getRange(templateRanges.written)
			.setFormula(templateFormulas.written(templateRanges))
		userSheet
			.getRange(templateRanges.avgWritten)
			.setFormula(templateFormulas.avgWritten(templateRanges))
		userSheet
			.getRange(templateRanges.mostWritten)
			.setFormula(templateFormulas.mostWritten(templateRanges))
		userSheet
			.getRange(templateRanges.avgRate)
			.setFormula(templateFormulas.avgRate(templateRanges))
		userSheet
			.getRange(templateRanges.highestRate)
			.setFormula(templateFormulas.highestRate(templateRanges))

		userSheet
			.getRange(templateRanges.goal)
			.setFormula(templateFormulas.goal(templateRanges, goalAmount))
		userSheet
			.getRange(templateRanges.completion)
			.setFormula(templateFormulas.completion(templateRanges, goalAmount))
		userSheet
			.getRange(templateRanges.daily)
			.setFormula(templateFormulas.daily(templateRanges, goalAmount))
		userSheet
			.getRange(templateRanges.weekly)
			.setFormula(templateFormulas.weekly(templateRanges, goalAmount))

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
		logMessage(logTypes.error, `An error has occured: ${err.stack}`, source)
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

	if (!goalTypes.includes(goalType) || typeof goalAmount !== "number") {
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

		const oldGoalType = getSheet(userSheetName)
			.getRange(templateRanges.metric)
			.getValue()
		logMessage(logTypes.debug, `oldGoalType: ${oldGoalType}`, source)

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
