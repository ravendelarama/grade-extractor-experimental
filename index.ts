import Exceljs from "exceljs";

interface Record {
    id: string;
    name: string;
    grade: number;
}

const spreadsheet = {
    path: './subject-grade-sample.xlsx',
    sheet: 1,
    headerRow: 1,
    selectedCols: [1, 2, 3]
};

async function extractor(filePath: string) {
    const workbook = new Exceljs.Workbook();
    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.getWorksheet(spreadsheet.sheet);

    const records = Array<Record>();

    worksheet?.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber > spreadsheet.headerRow) {
            const id = row.getCell(spreadsheet.selectedCols[0]).value as string;
            const name = row.getCell(spreadsheet.selectedCols[1]).value as string;
            const grade = row.getCell(spreadsheet.selectedCols[2]).value as number;

            records.push({
                id,
                name,
                grade
            })
        }
    })

    return records;
}

const results = await extractor(spreadsheet.path);

console.log("Subject Grades:");

results.forEach((result) => {
    console.log(`#${result.id} ${result.name}: ${result.grade}`);
});