const colors = require('colors');
const csv = require('csvtojson');
const path = require('path');
const excel = require('excel4node');

const getFilePath = (fileName) => path.join(__dirname, `${fileName}.csv`);

const readCSV = async (fileName) => {
    const filePath = getFilePath(fileName);
    const jsonArray = await csv().fromFile(filePath);
    return jsonArray;
};

const proccessCSV = (data, q, ssp) => {
    const uniqueSubs = [...new Set(data.map((item) => item.SUB))];
    const results = [];

    uniqueSubs.forEach((sub) => {
        const subData = data.filter((item) => item.SUB === sub);
        const uniqueMonths = [...new Set(subData.map((item) => item.MON))];

        uniqueMonths.forEach((month) => {
            const monthData = subData.filter((item) => item.MON === month);
            const baseline = monthData.find((item) => item.Periodo === 'Linea Base');
            const cortoPlazo = monthData.find((item) => item.Periodo === 'Corto Plazo');
            const medianoPlazo = monthData.find((item) => item.Periodo === 'mediano Plazo' || item.Periodo === 'Mediano Plazo');
            const largoPlazo = monthData.find((item) => item.Periodo === 'Largo Plazo');

            const result = {
                q,
                ssp,
                SUB: sub,
                SUB_Nombre: baseline.SUB_Nombre,
                MON: month,
                [baseline.Periodo]: baseline.FLOW_OUTcms,
                [cortoPlazo.Periodo]: cortoPlazo.FLOW_OUTcms,
                [medianoPlazo.Periodo]: medianoPlazo.FLOW_OUTcms,
                [largoPlazo.Periodo]: largoPlazo.FLOW_OUTcms,
                '% CP': ((cortoPlazo.FLOW_OUTcms - baseline.FLOW_OUTcms) / baseline.FLOW_OUTcms * 100).toFixed(1),
                '% MP': ((medianoPlazo.FLOW_OUTcms - baseline.FLOW_OUTcms) / baseline.FLOW_OUTcms * 100).toFixed(1),
                '% LP': ((largoPlazo.FLOW_OUTcms - baseline.FLOW_OUTcms) / baseline.FLOW_OUTcms * 100).toFixed(1)
            };

            results.push(result);
        });
    });

    return results;
};

const writeWorksheet = (headerStyle, worksheet, data) => {
    const sspUniqueSubs = [...new Set(data.map((item) => item.SUB))];
    let rowIndex = 1;

    sspUniqueSubs.forEach((sub) => {
        const subData = data.filter((item) => item.SUB === sub);
        const uniqueQ = [...new Set(subData.map((item) => item.q))];

        uniqueQ.forEach((q) => {
            const qData = subData.filter((item) => item.q === q);

            // Add header row before each qData
            worksheet.cell(rowIndex, 3).string('SUB').style(headerStyle);
            worksheet.cell(rowIndex, 4).string('SUB_Nombre').style(headerStyle);
            worksheet.cell(rowIndex, 5).string('MON').style(headerStyle);
            worksheet.cell(rowIndex, 6).string('Linea Base').style(headerStyle);
            worksheet.cell(rowIndex, 7).string('Corto Plazo').style(headerStyle);
            worksheet.cell(rowIndex, 8).string('Mediano Plazo').style(headerStyle);
            worksheet.cell(rowIndex, 9).string('Largo Plazo').style(headerStyle);
            worksheet.cell(rowIndex, 10).string('% CP').style(headerStyle);
            worksheet.cell(rowIndex, 11).string('% MP').style(headerStyle);
            worksheet.cell(rowIndex, 12).string('% LP').style(headerStyle);
            rowIndex++;

            qData.forEach((value) => {
                worksheet.cell(rowIndex, 1).string(value.q);
                worksheet.cell(rowIndex, 2).string(value.ssp);
                worksheet.cell(rowIndex, 3).string(value.SUB);
                worksheet.cell(rowIndex, 4).string(value.SUB_Nombre);
                worksheet.cell(rowIndex, 5).string(value.MON);
                worksheet.cell(rowIndex, 6).string(value['Linea Base']);
                worksheet.cell(rowIndex, 7).string(value['Corto Plazo']);
                worksheet.cell(rowIndex, 8).string(value['Mediano Plazo'] || value['mediano Plazo']);
                worksheet.cell(rowIndex, 9).string(value['Largo Plazo']);
                worksheet.cell(rowIndex, 10).string(value['% CP']);
                worksheet.cell(rowIndex, 11).string(value['% MP']);
                worksheet.cell(rowIndex, 12).string(value['% LP']);
                rowIndex++;
            });

            rowIndex++;
        });
    });
};

const processData = async () => {
    console.log('Processing data...'.rainbow);

    const scenarios = [
        { fileName: 'Datos_Promedios_Qmax_ssp119', q: 'QMX', ssp: 'SSP119' },
        { fileName: 'Datos_Promedios_Qmax_ssp226', q: 'QMX', ssp: 'SSP226' },
        { fileName: 'Datos_Promedios_Qmax_ssp245', q: 'QMX', ssp: 'SSP245' },
        { fileName: 'Datos_Promedios_Qmax_ssp585', q: 'QMX', ssp: 'SSP585' },
        { fileName: 'Datos_Promedios_Qmed_ssp119', q: 'QMED', ssp: 'SSP119' },
        { fileName: 'Datos_Promedios_Qmed_ssp226', q: 'QMED', ssp: 'SSP226' },
        { fileName: 'Datos_Promedios_Qmed_ssp245', q: 'QMED', ssp: 'SSP245' },
        { fileName: 'Datos_Promedios_Qmed_ssp585', q: 'QMED', ssp: 'SSP585' },
        { fileName: 'Datos_Promedios_Qmin_ssp119', q: 'QMIN', ssp: 'SSP119' },
        { fileName: 'Datos_Promedios_Qmin_ssp226', q: 'QMIN', ssp: 'SSP226' },
        { fileName: 'Datos_Promedios_Qmin_ssp245', q: 'QMIN', ssp: 'SSP245' },
        { fileName: 'Datos_Promedios_Qmin_ssp585', q: 'QMIN', ssp: 'SSP585' }
    ];

    const data = [];
    for (const scenario of scenarios) {
        console.log(`Reading data from CSV file ${scenario.fileName}...`.yellow);
        const csvData = await readCSV(scenario.fileName);
        const processedData = proccessCSV(csvData, scenario.q, scenario.ssp);
        data.push(...processedData);
    }

    console.log('Data processed successfully!'.green);

    const workbook = new excel.Workbook();
    const headerStyle = workbook.createStyle({
        font: {
            bold: true,
            size: 11
        }
    });

    const sspWorksheets = ['SSP119', 'SSP226', 'SSP245', 'SSP585'];
    sspWorksheets.forEach((ssp) => {
        const worksheet = workbook.addWorksheet(ssp);
        const sspData = data.filter((item) => item.ssp === ssp);
        writeWorksheet(headerStyle, worksheet, sspData);
    });

    workbook.write('data.xlsx');
};

processData();