/** @format */

//imports
const XLSX = require('xlsx')
const XML = require('xmlbuilder')
const FS = require('fs')

//variables
var now = new Date(Date.now())
var nowDate = [now.getFullYear(), now.getMonth() + 1, now.getDate()].join('-')
var nowTime = [now.getHours(), now.getMinutes()].join(':')
var storage = {}
var pathOutput = './output_' + nowDate + '_' + nowTime
var path = './Template_MdtExcel.xlsx'
//  '/Users/onaimann.capgemini/Library/Mobile Documents/com~apple~CloudDocs/Work.Capgemini/Projekte/MAN-ESM/BB_OM-Order-Management/MDT_GenericFieldMapping.xlsx'

//functions
//...

//---code logic---
//prepare outputfolder
if (!FS.existsSync(pathOutput)) {
  FS.mkdirSync(pathOutput)
}
//read whole excelfile
var wb = XLSX.readFile(path)
wb.SheetNames.forEach((wsName) => {
  if (wsName.toLowerCase().endsWith('__mdt')) {
    storage[wsName] = XLSX.utils.sheet_to_json(wb.Sheets[wsName], { raw: true, defval: null })
  }
})
//craft xml output
let deploymentNotes = 'CustomMetadata:\n'
for (let mdtName in storage) {
  for (let json of storage[mdtName]) {
    if (!json.MasterData && !json.DeveloperName) {
      continue
    }

    var xml = XML.create('CustomMetadata')
    xml.att('xmlns', 'http://soap.sforce.com/2006/04/metadata')
    xml.att('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance')
    xml.att('xmlns:xsd', 'http://www.w3.org/2001/XMLSchema')

    var lv1 = xml
    xml = lv1.ele('label', json.MasterLabel)
    xml = lv1.ele('protected', false)

    for (let key in json) {
      if (!key.startsWith('MasterLabel') && !key.startsWith('DeveloperName')) {
        let arrKey = key.split(',')
        var lv2 = lv1.ele('values')
        let tmp = lv2
        lv2 = tmp.ele('field', arrKey[0])
        if (json[key]) {
          lv2 = tmp.ele(
            'value',
            { 'xsi:type': 'xsd:' + arrKey[1] },
            arrKey[1] == 'boolean' ? JSON.parse(json[key].toLowerCase()) : json[key]
          )
        } else {
          lv2 = tmp.ele('value', { 'xsi:nil': 'true' }, '')
        }

        xml = lv2
      }
    }
    xml = xml.end({ pretty: true })

    //save xml content in file
    let targetPath = pathOutput + '/' + mdtName
    if (!FS.existsSync(targetPath)) {
      FS.mkdirSync(targetPath)
    }
    let arrName = [mdtName.split('__')[0], json.DeveloperName, 'md-meta', 'xml']
    let path = targetPath + '/' + arrName.join('.')
    FS.writeFile(path, xml, (err) => {
      if (err) throw err
    })
    deploymentNotes += '- ' + arrName[0] + '.' + arrName[1] + '\n'
  }
}
FS.writeFile(pathOutput + '/DeploymentNotes.txt', deploymentNotes, (err) => {
  if (err) throw err
})
console.log('\n\nmdt files saved in "' + pathOutput + '"')
