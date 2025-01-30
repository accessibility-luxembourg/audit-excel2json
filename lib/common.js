const getFieldVal = function (sheet, x, y, type) {
  if (sheet[x + y] !== undefined) {
    return sheet[x + y][type]
  } else {
    return ''
  }
}

function convertDateFormat (dateStr) {
  if (dateStr !== '') {
    // Split the input date string by '/'
    const [month, day, year] = dateStr.split('/')

    // Convert the year to a 4-digit format
    const fullYear = year.length === 2 ? '20' + year : year

    // Return the date in 'yyyy-mm-dd' format
    return `${fullYear}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`
  } else {
    return ''
  }
}

export { getFieldVal, convertDateFormat }
