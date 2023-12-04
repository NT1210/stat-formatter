const XlsxPopulate = require('xlsx-populate');
const fs = require("fs")
const path = require("path");

function extract(file){
    let fullPath = "./" + file

    XlsxPopulate.fromFileAsync(fullPath)
    .then(workbook => {
        let sheet = workbook.sheet("Sheet1")

        const targetList = [
            "дата выпуска",
            "наименование отправителя",
            "адрес отправителя",
            "код страны отправителя",
            "наименование получателя",
            "код страны получателя",
            "адрес получателя",
            "наименование контрактодержателя",
            "код страны контрактодержателя",
            "адрес контрактодержателя",
            "страна происхождения",
            "наименование и характеристики товаров",
            "фирма-изготовитель",
            "вес нетто",
            "статистическая стоимость",
            "usdkg",
            "код товара по тн"
        ]

        let tempObj = {}
        
        for(let i=1; i<101; i++){
            let title = sheet.row(1).cell(i).value()
            
            if(title === undefined) break
            
            for(let list of targetList){
                if((title.toLowerCase()).includes(list)) {
                    let tempArr = []

                    for(let j=2; j<10000; j++){

                        let val = sheet.row(j).cell(i).value()
                        // if(val === undefined) break
                        tempArr.push(val)
                    }

                    tempObj[title] = tempArr
                    tempArr = []
                }
            }   
        }

        const propForDelete = 'G31_13 (Страна происхождения)'
        delete tempObj[propForDelete]
        
        tempObj["fileName"] = file
        
        return tempObj
            
    }).then(tempObj => {
        XlsxPopulate.fromBlankAsync()
            .then(workbook2 => {
                let newfileName = (tempObj["fileName"]).split("\\").slice(-1)[0]

                delete tempObj.fileName
                let sheet = workbook2.sheet("Sheet1")
                let titles = Object.keys(tempObj)
            
                for(let title of titles){
                    let idxOfTitle 
                    let arrForOutput = tempObj[title]

                    if( (title.toLowerCase()).includes("дата выпуска") ) idxOfTitle=1
                    if( (title.toLowerCase()).includes("наименование отправителя") ) idxOfTitle=2
                    if( (title.toLowerCase()).includes("адрес отправителя") ) idxOfTitle=3
                    if( (title.toLowerCase()).includes("код страны отправителя") ) idxOfTitle=4
                    if( (title.toLowerCase()).includes("наименование получателя") ) idxOfTitle=5
                    if( (title.toLowerCase()).includes("код страны получателя") ) idxOfTitle=6
                    if( (title.toLowerCase()).includes("адрес получателя") ) idxOfTitle=7
                    if( (title.toLowerCase()).includes("наименование контрактодержателя") ) idxOfTitle=8
                    if( (title.toLowerCase()).includes("код страны контрактодержателя") ) idxOfTitle=9
                    if( (title.toLowerCase()).includes("адрес контрактодержателя") ) idxOfTitle=10
                    if( (title.toLowerCase()).includes("страна происхождения") ) idxOfTitle=11
                    if( (title.toLowerCase()).includes("наименование и характеристики товаров") ) idxOfTitle=12
                    if( (title.toLowerCase()).includes("фирма-изготовитель") ) idxOfTitle=13
                    if( (title.toLowerCase()).includes("вес нетто") ) idxOfTitle=14
                    if( (title.toLowerCase()).includes("статистическая стоимость") ) idxOfTitle=15
                    if( (title.toLowerCase()).includes("usdkg") ) idxOfTitle=16
                    if( (title.toLowerCase()).includes("код товара по тн") ) {
                        idxOfTitle=17
                        tempObj[title] = tempObj[title].map(ele => {
                            return parseInt(ele)
                        })
                    }
              
                    sheet.row(1).cell(idxOfTitle).value(title)

                    arrForOutput.forEach((ele, idx) => {
                        sheet.row(idx+2).cell(idxOfTitle).value(ele)
                    })
                }

                let beforeRename = newfileName.split(".")
                let beforeRename2 = `${beforeRename[0]}-extracted.${beforeRename[1]}`

                const outputFilePath = `./output/${beforeRename2}`

                return workbook2.toFileAsync(outputFilePath)
            })
    })
}

function main() {
    const files = fs.readdirSync("./input")
    
    for(let file of files){
        let year = file.includes("2023") ? 2023 : 2022
        let relPath = year === 2023 ? path.join("2023", "original", file) : path.join("2022", "original", file)
        extract(relPath)
    }
}


main()