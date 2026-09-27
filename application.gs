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

		if (response !== "edit") {
			const totalSheets = spreadsheet.getNumSheets()
			spreadsheet.insertSheet(userSheetName, totalSheets, {
				template: getSheet(sheetNames.template),
			})
		}

		const userSheet = getSheet(userSheetName)

		if (response === "edit") {
			// Sheet already exists: swap the old metric wording for the new one, everywhere it appears
			const oldGoalType = userSheet.getRange(templateRanges.metric).getValue()
			if (oldGoalType !== goalType) {
				userSheet
					.createTextFinder(oldGoalType)
					.matchEntireCell(false)
					.replaceAllWith(goalType)
			}
		} else {
			// New sheet: swap the [[METRIC]] placeholder for the goal type they picked
			userSheet
				.createTextFinder("[[METRIC]]")
				.matchEntireCell(false)
				.replaceAllWith(goalType)
		}

		// Store the goal amount as a real value; the template's own formulas read
		// E4 directly, so nothing else needs to be set here
		userSheet.getRange(templateRanges.goalAmount).setValue(goalAmount)

		const nameRange = userSheet.getRange(templateRanges.name)
		nameRange.setValue(name)

		if (response !== "edit") {
			nameRange.protect().setWarningOnly(true)
			userSheet.getRange(templateRanges.metric).protect().setWarningOnly(true)
		}

		userSheet.setTabColor(null)
		applicationSheet.getRange(applicationRange.full).clearContent()

		logMessage(
			logTypes.info,
			`Finished ${response == "edit" ? "editing" : "creating"} sheet according to the application.`,
			source,
		)
		SpreadsheetApp.getUi().alert(
			`Finished ${response == "edit" ? "editing" : "creating"} sheet according to the application.`,
		)
	} catch (err) {
		logMessage(logTypes.error, `An error has occurred: ${err.stack}`, source)
		SpreadsheetApp.getUi().alert(
			`An error has occurred, please contact the script creator for help: ${err.stack}`,
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

function checkExistingApplication(userSheetName) {
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

		return "edit"
	}

	return true
}
