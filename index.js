const express = require('express');
const exceljs = require("exceljs");
const app = express();
const port = 3000;
const data = require("./data.json")


//Data JSON COMO EJEMPLO, // TODO: Remplazar para la migracion
app.get("/", (request, response) =>{
    const workbook = exportData(data);
    response.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    response.setHeader(
        "Content-Disposition",
        "attachment; filename=" + "data.xlsx"
    );

    return workbook.xlsx.write(response).then(function(){
        response.status(200).end();
    });
});

app.listen(port, () =>{
    console.log(`Listening at port ${port}`);
});

//Name sheet
const exportData = (data) =>{
    let workbook = new exceljs.Workbook()
    let sheet = workbook.addWorksheet('A-PA');
    let colums = data.reduce((acc, obj) => acc = Object.getOwnPropertyNames(obj), [])
    

    sheet.columns = colums.map((excel) =>{
        return {header: excel, key: excel, width: 20}
    });
    sheet.addRows(data);
    return workbook;
};