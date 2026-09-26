let automaticUpdate = true

function countStat(stat = "sheet") {
	const allSheets = spreadsheet.getSheets()

	let count = 0

	allSheets.forEach((sheet) => {
		// Skip non-user sheets
		if (
			Object.values(sheetNames).includes(sheet.getName()) ||
			sheet.getName() == "Bob"
		)
			return

		if (stat == "sheet") count++

		if (stat == "task") {
			const headerValue = sheet.getRange("B2").getValue()

			if (headerValue === "Your accomplishments") {
				// Task goal sheet
				if (stat == "task") count += Number(sheet.getRange("D6").getValue()) || 0
			}
		}
	})

	return count
}

function averageStat(stat = "completion") {
	const allSheets = spreadsheet.getSheets()

	let count = 0
	let sum = 0

	allSheets.forEach((sheet) => {
		// Skip non-user sheets
		if (
			Object.values(sheetNames).includes(sheet.getName()) ||
			sheet.getName() == "Bob"
		)
			return

		const headerValue = sheet.getRange("B2").getValue()

		if (stat == "completion") {
			const completionRange =
				headerValue === "Your accomplishments" ? "D10" : "D20"
			count++
			sum +=
				parseFloat(
					String(sheet.getRange(completionRange).getValue()).replace("%", ""),
				) || 0
		}
	})

	return count > 0 ? sum / count : 0
}
