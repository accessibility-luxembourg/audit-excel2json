// ce script écrit en nodeJS permet de convertir un audit d'accessibilité (grille Excel)en fichier au format JSON
// fichiers requis :
// allNames.json (notre inventaire de sites et apps)
// common.js (bibliothèque qui comprend des fonctions, variables et tableaux spécifiques au projet AccessLux)
// éléments du dossier criteria (les référentiels)
// audit_step1_convert.sh pour un traitement en lot
//
// important : ajouter une ligne "Plateforme :" pour préciser, dans les audits RAAM, s'il s'agit d'un audit iOS ou Android

// chargement des bibliothèques requises par le projet : XLSX, fs et common.js.
import * as XLSX from 'xlsx/xlsx.mjs'
import * as fs from 'fs'
import * as lib from './lib/common.js'


XLSX.set_fs(fs)

// option 1 : récupérer le nom de fichier passé via le script audit-full_step1_convert.sh (utile pour un traitement par lot)
const workbook = XLSX.readFile(process.argv[2])

// option 2 : définir ici le nom de fichier
// const workbook = XLSX.readFile('./file_to_read.xlsx')

// sélection de l'onglet Echantillon ou Échantillon (variantes selon les grilles)
let sample = workbook.Sheets.Echantillon
let sampleName = 'Echantillon'
if (sample === undefined) {
  sample = workbook.Sheets.Échantillon
  sampleName = 'Échantillon'
}

// fonctions spécifiques au script
// lecture de la date d'un audit simplifié
function convertDateFormatSimple (dateStr) {
  let nDate = dateStr.split(': ').pop()
  nDate = nDate.split(' ')
  const months = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre']
  const monthNr = months.indexOf(nDate[1].toLowerCase())
  return (nDate[2] + '-' + (monthNr + 1).toString().padStart(2, '0') + '-' + nDate[0].padStart(2, '0'))
}

// lecture de la dérogation pour un critère, dans une page
function getExemption (refPage, col, line) {
  const str = lib.getFieldVal(workbook.Sheets[refPage], col, line, 'v').toLowerCase().trim()
  if (str === 'd') {
    return true
  } else if (str === 'e') {
    return true
  } else if (str === '') {
    return false
  } else if (str === 'n') {
    return false
  } else {
    console.error('Unknown exemption value:', str, 'page:', refPage, 'col:', col, 'line:', line)
    return false
  }
}

// lecture du statut d'un critère, pour une page
function getStatusName (refPage, col, line, exemption) {
  const str = lib.getFieldVal(workbook.Sheets[refPage], col, line, 'v').toUpperCase().trim()
  if (!['C', 'NC', 'NA', 'NT'].includes(str)) {
    console.error('Unknown status value:', str, 'page:', refPage, 'col:', col, 'line:', line, 'exemption:', exemption)
  }
  return str
}

// obtenir la valeur d'un champ dans une cellule précise
function getTextValue (refPage, col, line) {
  const str = lib.getFieldVal(workbook.Sheets[refPage], col, line, 'v')
  if (str === undefined) {
    console.error('undefined, cannot be cleaned up', 'refPage:', refPage, 'col:', col, 'line:', line)
    return undefined
  } else {
    return str.toString().trim()
  }
}

// définition des référentiels
const refCriteria = {}
refCriteria.Complet = {}
refCriteria.Complet.RGAA = JSON.parse(fs.readFileSync('./data/criteria/rgaa412/RGAA-4.1.2.json')).topics.flatMap(e => e.criteria.map(f => e.number + '.' + f.criterium.number))
refCriteria.Complet.RAWeb = JSON.parse(fs.readFileSync('./data/criteria/raweb1/raweb.json')).topics.flatMap(e => e.criteria.map(f => e.number + '.' + f.criterium.number))
refCriteria['Simplifié'] = {}
refCriteria['Simplifié'].RGAA = Object.keys(JSON.parse(fs.readFileSync('./data/criteria/simple1/levels.json')))
refCriteria['Simplifié'].RAWeb = Object.keys(JSON.parse(fs.readFileSync('./data/criteria/simple1/levels.json')))
refCriteria.Mobile = {}
refCriteria.Mobile.RAAM = JSON.parse(fs.readFileSync('./data/criteria/raam1/raam1.json')).topics.flatMap(e => e.criteria.map(f => e.number + '.' + f.criterium.number))

// obtenir le critère indiqué dans une cellule précise
function getCriterionNum (refPage, col, line, reference, controlType) {
  const val = getTextValue(refPage, col, line)
  if (refCriteria[controlType] !== undefined && refCriteria[controlType][reference] !== undefined) {
    if (!refCriteria[controlType][reference].includes(val)) {
      console.error('Unknown criterion number:', val, 'at page:', refPage, 'col:', col, 'line:', line, 'reference:', reference, 'control type:', controlType)
    }
  } else {
    console.error('Control type / Reference not found')
    process.exit(1)
  }
  return val
}

// vérification que l'ensemble des critères a été testé
function checkAssessments (assessments, refPage, reference, controlType) {
  if (refCriteria[controlType] !== undefined && refCriteria[controlType][reference] !== undefined) {
    const criteriaInAssessment = new Set(assessments.map(e => e.criterion.number))
    const missingCriteria = refCriteria[controlType][reference].filter(x => !criteriaInAssessment.has(x))
    if (missingCriteria.length !== 0) {
      console.error('Missing criteria in assessment on page:', refPage, 'reference:', reference, 'controlType:', controlType, missingCriteria)
    }
  } else {
    console.error('Control type / Reference not found')
    process.exit(1)
  }
}

// attribution d'un statut pour l'ensembles des pages d'un critère
function mergeAssessments (asst1, asst2) {
  const newAsst = {}
  newAsst.criterion = asst1.criterion
  const statuses = [asst1.status.name, asst2.status.name]
  newAsst.status = {}
  if (statuses.filter(x => x === 'NC').length > 0) {
    newAsst.status.name = 'NC'
  } else if (statuses.filter(x => x === 'C').length > 0) {
    newAsst.status.name = 'C'
  } else if (statuses.filter(x => x === 'NA').length > 0) {
    newAsst.status.name = 'NA'
  } else {
    newAsst.status.name = 'NT'
  }
  newAsst.exemption = asst1.exemption || asst2.exemption
  newAsst.exemption_comment = (asst1.exemption_comment + '\r\n' + asst2.exemption_comment).trim()
  newAsst.changes_to_do = (asst1.changes_to_do + '\r\n' + asst2.changes_to_do).trim()
  newAsst.patches_follow_up = ''
  return newAsst
}

// début des opérations
// définition d'une structure json contenant toutes les données d'un audit
const audit = {}

// informations générales
const auditInfos = []

const fieldsToRetrieve = ['Type', 'Date', 'Entreprise', 'Contexte', 'Site', 'Plateforme', 'Application', 'Référentiel', 'Version référentiel']

// recherche des informations générales dans les premières lignes du premier onglet de la grille
for (let l = 1; l < 20; l++) {
  for (let f = 0; f < fieldsToRetrieve.length; f++) {
    if (lib.getFieldVal(sample, 'A', l, 'v').indexOf(fieldsToRetrieve[f]) > -1) { auditInfos[fieldsToRetrieve[f]] = lib.getFieldVal(sample, 'B', l, 'w') }
  }
}

// définition du type d'audit : Complet, Mobile, Simplifié et de la plateforme : Web, iOS, Android
audit.control_type = { name: 'Complet' }
if (lib.getFieldVal(sample, 'A', 1, 'v').indexOf('SIMPLIFIÉ') > -1) { audit.control_type = { name: 'Simplifié' } }

if (Object.hasOwn(auditInfos, 'Plateforme')) {
  audit.control_type = { name: 'Mobile' }
} else {
  auditInfos.Plateforme = 'Web'
}

// obtention de la date
if (Object.hasOwn(auditInfos, 'Date')) {
  if (auditInfos.Date === '') { // cas des audits simplifiés où la date est dans le même champ
    for (let l = 1; l < 10; l++) {
      if (lib.getFieldVal(sample, 'A', l, 'v').indexOf('Date') > -1) { auditInfos.Date = convertDateFormatSimple(lib.getFieldVal(sample, 'A', l, 'w')) }
    }
  } else {
    auditInfos.Date = lib.convertDateFormat(auditInfos.Date)
  }
}

// obtention des autres informations générales
audit.auditor = 'Non renseigné'
audit.context = 'Non renseigné'
if (Object.hasOwn(auditInfos, 'Entreprise')) {
  audit.auditor = auditInfos.Entreprise
}
if (Object.hasOwn(auditInfos, 'Contexte') && audit.context !== '') {
  audit.context = auditInfos.Contexte
}
audit.audited_at = auditInfos.Date
audit.platform = { name: auditInfos.Plateforme }
audit.company = auditInfos.Entreprise
if (audit.company === '' || audit.company === undefined) { audit.company = 'Non renseigné' }
audit.inventory = {}
audit.inventory.name = auditInfos.Site
if (audit.inventory.name === undefined || audit.inventory.name === '') {
  audit.inventory.name = auditInfos.Application
}

// vérification que le nom est dans l'inventaire
if (fs.existsSync('./out/inventory/allNames.json')) {
  const allNames = JSON.parse(fs.readFileSync('./out/inventory/allNames.json'))
  if (!allNames.includes(audit.inventory.name)) {
    console.error('Error: name not in the inventory:', audit.inventory.name, 'file:', process.argv[2])
  }
}

// détermination du référentiel qui s'applique
if (Object.hasOwn(auditInfos, 'Référentiel') && Object.hasOwn(auditInfos, 'Version référentiel')) {
  audit.audit_reference = { name: auditInfos['Référentiel'], version: auditInfos['Version référentiel'] }
} else {
  audit.audit_reference = { name: 'RGAA', version: '4.1.2' } // audits simplifiés
  if (audit.control_type.name === 'Mobile') {
    audit.audit_reference = { name: 'RAAM', version: '1.1' } // 1.1 à partir de 2024
  }
  if (audit.control_type.name === 'Complet') {
    for (let l = 1; l < 10; l++) {
      if (lib.getFieldVal(sample, 'A', l, 'v').toLowerCase().indexOf('raweb') > -1) { audit.audit_reference = { name: 'RAWeb', version: '1' } }
      if (lib.getFieldVal(sample, 'B', l, 'w').toLowerCase().indexOf('raweb') > -1) { audit.audit_reference = { name: 'RAWeb', version: '1' } }
    }
  }
  if (audit.control_type.name === 'Simplifié' && audit.audited_at.indexOf('2024') > -1) {
    audit.audit_reference = { name: 'RAWeb', version: '1' }
  }
}

audit.assessed_level = {}
audit.assessed_level.name = 'AA' // valeur commune à tous les audits de conformité

audit.pages = []

// échantillon de pages : à partir de quelle ligne commence le listing ?
let sampleStartLine
for (let l = 1; l < 20; l++) {
  if (lib.getFieldVal(sample, 'A', l, 'v').toLowerCase().indexOf('p01') > -1) { sampleStartLine = l }
  if (lib.getFieldVal(sample, 'A', l, 'v').toLowerCase().indexOf('e01') > -1) { sampleStartLine = l }
}

sampleStartLine--
let uri = ''
let refPage = ''
let titlePage = ''
let number = 1
uri = (audit.control_type.name === 'Mobile') ? getTextValue(sampleName, 'B', sampleStartLine + number) : getTextValue(sampleName, 'C', sampleStartLine + number)
refPage = getTextValue(sampleName, 'A', sampleStartLine + number)
titlePage = getTextValue(sampleName, 'B', sampleStartLine + number)

// récupération des notes
do {
  let assessments = {}

  if (workbook.Sheets[refPage] !== undefined) {
    for (let l = 4; l < 170; l++) {
      const assessment = {}
      assessment.criterion = {}
      switch (audit.control_type.name) {
        case 'Complet':
          // ne pas importer les statuts AAA
          if (['A', 'AA'].includes(lib.getFieldVal(workbook.Sheets[refPage], 'C', l, 'v'))) {
            assessment.criterion.number = getCriterionNum(refPage, 'B', l, audit.audit_reference.name, audit.control_type.name)
            assessment.status = {}
            if (lib.getFieldVal(workbook.Sheets[refPage], 'F', 3, 'v') === 'Récurrent') {
              assessment.exemption = getExemption(refPage, 'G', l)
              assessment.exemption_comment = getTextValue(refPage, 'I', l)
              assessment.changes_to_do = getTextValue(refPage, 'H', l)
              assessment.patches_follow_up = ''
            } else { // sur certains audits complets, les dérogations sont en colonne F
              assessment.exemption = getExemption(refPage, 'F', l)
              assessment.exemption_comment = getTextValue(refPage, 'H', l)
              assessment.changes_to_do = getTextValue(refPage, 'G', l)
              assessment.patches_follow_up = ''
            }
            assessment.status.name = getStatusName(refPage, 'E', l, assessment.exemption)
          }
          break

        case 'Mobile':
          if (lib.getFieldVal(workbook.Sheets[refPage], 'C', 3, 'v') === 'Critère') {
            // ne pas importer les critères AAA
            if (['A', 'AA'].includes(lib.getFieldVal(workbook.Sheets[refPage], 'D', l, 'v'))) {
              assessment.criterion.number = getCriterionNum(refPage, 'C', l, audit.audit_reference.name, audit.control_type.name)
              assessment.status = {}
              assessment.exemption = getExemption(refPage, 'G', l)
              assessment.status.name = getStatusName(refPage, 'F', l, assessment.exemption)
              assessment.exemption_comment = getTextValue(refPage, 'I', l)
              assessment.changes_to_do = getTextValue(refPage, 'H', l)
              assessment.patches_follow_up = ''
            }
          } else { // sur certains audits RAAM, les statuts sont en colonne C
            // ne pas importer les critères AAA
            if (['A', 'AA'].includes(lib.getFieldVal(workbook.Sheets[refPage], 'C', l, 'v'))) {
              assessment.criterion.number = getCriterionNum(refPage, 'B', l, audit.audit_reference.name, audit.control_type.name)
              assessment.status = {}
              assessment.exemption = getExemption(refPage, 'F', l)
              assessment.status.name = getStatusName(refPage, 'E', l, assessment.exemption)
              assessment.exemption_comment = getTextValue(refPage, 'H', l)
              assessment.changes_to_do = getTextValue(refPage, 'G', l)
              assessment.patches_follow_up = ''
            }
          }
          break

        case 'Simplifié':
          if (lib.getFieldVal(workbook.Sheets[refPage], 'B', l, 'v') !== '') {
            assessment.criterion.number = getCriterionNum(refPage, 'B', l, audit.audit_reference.name, audit.control_type.name)
            assessment.status = {}
            assessment.exemption = getExemption(refPage, 'E', l)
            assessment.status.name = getStatusName(refPage, 'D', l, assessment.exemption)
            assessment.exemption_comment = getTextValue(refPage, 'G', l)
            assessment.changes_to_do = getTextValue(refPage, 'F', l)
            assessment.patches_follow_up = ''
          }
          break
      }
      if (assessment.status !== undefined && assessment.status.name !== undefined && assessment.status.name !== '') {
        if (assessments[assessment.criterion.number] !== undefined) {
          assessments[assessment.criterion.number] = mergeAssessments(assessments[assessment.criterion.number], assessment)
        } else {
          assessments[assessment.criterion.number] = assessment
        }
      }
    }
  }
  if ((uri !== undefined && uri !== '') || (titlePage !== undefined && titlePage !== '')) {
    if (uri === undefined || uri === '') { uri = 'non-defini' }
    assessments = Object.values(assessments)
    checkAssessments(assessments, refPage, audit.audit_reference.name, audit.control_type.name)
    audit.pages.push({ number, uri, assessments })
  }

  number++
  uri = (audit.control_type.name === 'Mobile') ? getTextValue(sampleName, 'B', sampleStartLine + number) : getTextValue(sampleName, 'C', sampleStartLine + number)
  refPage = getTextValue(sampleName, 'A', sampleStartLine + number)
  titlePage = getTextValue(sampleName, 'B', sampleStartLine + number)
} while ((uri !== undefined && uri !== '') || (titlePage !== undefined && titlePage !== ''))

// sortie du JSON sur la console (possibilité de le stocker dans un fichier via le script complet-convert.sh)
console.log(JSON.stringify(audit, null, 2))
