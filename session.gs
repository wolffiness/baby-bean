function addSessionToLogs() {
	const source = "addSessionToLogs"

	try {
		const userSheet = spreadsheet.getActiveSheet()
		const logSheet = getSheet(sheetNames.logs)

		// Move userSheet to front of spreadsheet
		const numSheets = Object.keys(sheetNames).length
		userSheet.activate()
		spreadsheet.moveActiveSheet(numSheets + 1)
		logMessage(
			logTypes.debug,
			`Moved ${userSheet.getName()}'s sheet to position ${numSheets + 1}.`,
			source,
		)

		// Get form data from the Dashboard sheet
		const name = userSheet.getRange(logSessionRange.name).getValue()
		const date = userSheet.getRange(logSessionRange.date).getValue()
		const timeStart = userSheet.getRange(logSessionRange.timeStart).getValue()
		const timeEnd = userSheet.getRange(logSessionRange.timeEnd).getValue()
		const countStart = userSheet.getRange(logSessionRange.countStart).getValue()
		const countEnd = userSheet.getRange(logSessionRange.countEnd).getValue()

		// Validate the inputs
		if (
			!name ||
			!date ||
			!timeStart ||
			!timeEnd ||
			(!countStart && countStart !== 0) ||
			!countEnd
		) {
			logMessage(
				logTypes.warning,
				"Please fill in all the required fields: Name, Date, Start and End times, Start and End word count.",
				source,
			)
			logMessage(
				logTypes.debug,
				`name: ${name}. date: ${date}. timeStart: ${timeStart}. timeEnd: ${timeEnd}. countStart: ${countStart}. countEnd: ${countEnd}`,
				source,
			)
			SpreadsheetApp.getUi().alert(
				"Please fill in all the required fields: Name, Date, Start and End times, Start and End word count.",
			)
			return
		}

		if (
			typeof date !== "object" ||
			typeof timeStart !== "object" ||
			typeof timeEnd !== "object"
		) {
			logMessage(
				logTypes.warning,
				`Please make sure the date, start, and end times are all in a valid format.`,
				source,
			)
			logMessage(
				logTypes.debug,
				`date: ${date}. timeStart: ${timeStart}. timeEnd: ${timeEnd}`,
				source,
			)
			SpreadsheetApp.getUi().alert(
				`Please make sure the date, start, and end times are all in a valid format.`,
			)
			return
		}

		if (typeof countStart !== "number" || typeof countEnd !== "number") {
			logMessage(
				logTypes.warning,
				`Please make sure the start and end count are numbers.`,
				source,
			)
			logMessage(
				logTypes.debug,
				`countStart: ${countStart}. countEnd: ${countEnd}`,
				source,
			)
			SpreadsheetApp.getUi().alert(
				`Please make sure the start and end count are numbers.`,
			)
			return
		}

		// Append the new prompt to the Logs sheet
		const newRow = [
			name,
			date,
			timeStart,
			timeEnd,
			countStart,
			countEnd,
			"",
			"",
			"",
		]
		Logger.log(newRow)
		logSheet.appendRow(newRow)

		const lastRow = logSheet.getLastRow()
		logSheet.getRange(lastRow, 7).setFormula(`=F${lastRow}-E${lastRow}`) // Words written
		logSheet.getRange(lastRow, 8).setFormula(`=D${lastRow}-C${lastRow}`) // Session duration
		logSheet
			.getRange(lastRow, 9)
			.setFormula(`=ROUND(G${lastRow}/(H${lastRow}*1440),0)`) // WPM

		// Clear the input fields in the Dashboard and confirm success
		userSheet.getRange(logSessionRange.countStart).setValue(countEnd)
		userSheet.getRange(logSessionRange.countEnd).clearContent()
		userSheet.getRange(logSessionRange.timeStart).clearContent()
		userSheet.getRange(logSessionRange.timeEnd).clearContent()

		createTrigger(statCalculations.completion.trigger_function, 5)
		createTrigger(statCalculations.tasks_completed.trigger_function, 5)

		logMessage(logTypes.info, "Session has been saved.", source)
		SpreadsheetApp.getUi().alert("Session has been saved.")
	} catch (error) {
		logMessage(logTypes.error, `An error has occured: ${error.stack}`, source)
		SpreadsheetApp.getUi().alert(
			`An error has occured, please contact the script creator for help: ${error.stack}`,
		)
	}
}
