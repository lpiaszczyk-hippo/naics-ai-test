const exceljs = require('exceljs');

async function getDataTable(limit) {
    let workbook = new exceljs.Workbook();
    return workbook.xlsx.readFile('input.xlsx').then(function() {
        let worksheet = workbook.getWorksheet(1);
        worksheet.spliceRows(1, 1); //remove header

        let data = [];
        worksheet.eachRow(function(row, rowNumber) {
            
                let rowValues = [];
                row.eachCell(function(cell) {
                    rowValues.push(cell.value);
                });
                data.push(rowValues);
            
        });
        return data.map(value => ({ value, sort: Math.random() }))
        .sort((a, b) => a.sort - b.sort)
        .map(({ value }) => value)
        .splice(0, limit); //shuffle and return limit rows
    });
}

async function fetchNaics(url, version) {
    url = `http://localhost:8081/api/v1/naics?version=${version}&url=${url}`
    console.log('fetching url: ' + url);
    const start = performance.now();
    const response = await fetch(url);
    const jsonResponse = await response.json();
    const end = performance.now();
    return {codes: jsonResponse.codes, time: end - start};
}

function writeResults(results, version) { 
    const workbook = new exceljs.Workbook();
    const worksheet = workbook.addWorksheet('Results');
    worksheet.columns = [
        { header: 'URL', key: 'url', width: 50 },
        { header: 'Expected', key: 'expected', width: 10 },
        { header: 'Result', key: 'result', width: 10 },
        { header: 'Expected_code', key: 'expected_code', width: 10 },
        { header: 'Expected_desc', key: 'expected_desc', width: 10 },
        { header: 'Result_code', key: 'result_code', width: 10 },
        { header: 'Result_desc', key: 'result_desc', width: 10 },
        { header: 'Code_hit', key: 'code_hit', width: 10 },
        { header: 'Desc_hit', key: 'desc_hit', width: 10 },
        { header: 'Time', key: 'time', width: 10 },
    ];
    results.forEach(result => {
        worksheet.addRow({url: result.url, 
            expected: result.expected, 
            result: result.result, 
            expected_code: result.expected.value, 
            result_code: result.result.filter(item => item.confidence >= 0.8).map(item => item.code), 
            expected_desc: result.expected.description, 
            result_desc: result.result.filter(item => item.confidence >= 0.8).map(item => item.description), 
            code_hit: result.result.filter(item => item.code == result.expected.value).length >= 1 ? 'yes' : 'no',
            desc_hit: result.result.filter(item => item.description == result.expected.description).length >= 1 ? 'yes' : 'no',
            time: result.time});
    });
    workbook.xlsx.writeFile(`results_${version}.xlsx`);
}

async function run(size) {
    if(size == undefined || size == null || size == "" || size == 0) {
        throw new Error('specify sample size');
    }
    let table = await getDataTable(size);
    console.log(table);
    const reults = [];
    for (item of table) {
        const rawUrl = item[2];
        const encodedUrl = encodeURIComponent(item[2]);
        const naics = await fetchNaics(encodedUrl, version);
        reults.push({url: rawUrl, expected: JSON.parse(item[3])[0], result: naics.codes, time: naics.time});
    }
    // table.forEach(item => {
    //     console.log(JSON.parse(item[3])[0]) //get naics
    // });
    console.log(reults);
    writeResults(reults, version);
}

const sample = process.argv[2];
const version = process.argv[3];
run(sample);
