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

		if (stat == "completion") {
			count++
			sum +=
				parseFloat(
					String(
						sheet.getRange(templateRanges.completion).getValue(),
					).replace("%", ""),
				) || 0
		}
	})

	return count > 0 ? sum / count : 0
}