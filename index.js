const _ = require('lodash')
const fs = require('fs')
const path = require('path')
const XLSX = require('xlsx')

const dbFile = path.resolve(__dirname, 'vpmsdb_6.1.xlsx')
const { Sheets } = XLSX.readFile(dbFile, { raw: true })

const rawCodebook = _.get(Sheets, 'Codebook')
const jsonCodebookRows = XLSX.utils.sheet_to_json(rawCodebook, { header: 1, defval: null })

const codebookMappingExtractor = /([^=]+)\s+=\s(\d*)/

const codebookByLabel = jsonCodebookRows.slice(1).reduce((memo, row) => {
  const nullFilteredRow = row.filter(v => v !== null)
  if (_.isEmpty(nullFilteredRow)) return memo
  if (nullFilteredRow.length < 3) return memo

  // Attempt int/enum mapping
  const { label, description, ...codes } = nullFilteredRow.reduce((memo, cell, index) => {
    if (index === 0) return { ...memo, label: cell }

    const mapping = codebookMappingExtractor.exec(cell)
    if (!mapping) {
      if (index === 1) return { ...memo, description: cell }
      return { ...memo, [index - 2]: cell }
    }

    const label = _.get(mapping, 1)
    const number = _.get(mapping, 2)
    if (!label || !number) return memo

    return { ...memo, [number]: label }
  }, {})

  return { ...memo, [label.trim()]: { description, codes } }
}, {})

const rawIncidents = _.get(Sheets, 'Full Database')
const jsonIncidentRows = XLSX.utils.sheet_to_json(rawIncidents, { range: 1, defval: null })

console.log(jsonIncidentRows.map(incident => {
  return Object.entries(incident).reduce((memo, [ label, value ]) => {
    return { ...memo, [label.trim()]: _.get(codebookByLabel, `${label.trim()}.codes.${value}`, value) }
  }, {})
}))
