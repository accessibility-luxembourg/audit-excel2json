# audit-excel-to-json
This script converts accessibility audits in Excel format to files in json format. The goal is to have standardised machine readable files for audits, independent of the presentation in the Excel template.


# Installation

Install node.js. This script has been tested on Node.js v20.6.

Install the required packages with the following command:

```
npm install
```


# Usage

You can convert one file at a time with the following command:

```
node ./audit_convert.js yourAuditFile.xlsx > yourAuditFile.json
```

If executing this command on Windows, start the command with `node.exe` instead of `node`.


If you would like to convert multiple files in a given folder, you can use the associated shell script:
```
./convert_all.sh "folder"
```

This script will create a subfolder named "json" in which all the json files will be stored.

## License
This software is (c) [Information and press service](https://sip.gouvernement.lu/en.html) of the luxembourgish government and licensed under the MIT license.
