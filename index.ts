import ExcelJS from 'exceljs';

const DATA_START_ROW = 6;

interface ProductData {
    id: number;
    name: string;
    price: number;
    threeYearAgoSales: number;
    prevPrevYearSales: number;
    prevYearSales: number;
    sales: number[];
}

async function modifyExcel(): Promise<void> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('template.xlsx');
    const worksheet = workbook.getWorksheet(1);

    if (!worksheet) {
        throw new Error('Worksheet not found');
    }

    const currentDate = new Date();
    const currentMonth = currentDate.getMonth() + 1;
    const currentYear = currentDate.getFullYear();
    const fiscalYear = currentMonth >= 4 ? currentYear : currentYear - 1;

    // Set fiscal years in D4, E4, F4 cells
    const fisicalYearCells = ['D4', 'E4', 'F4'];
    fisicalYearCells.forEach((cell, index) => {
        const prefix = index === 0 ? '～' : '';
        worksheet.getCell(cell).value = `${prefix}${fiscalYear - (fisicalYearCells.length - index)}`;
    });

    // Set months in G4, I4, K4, M4 cells
    const months = ['04', '05', '06', '07'];
    const cells = ['G4', 'I4', 'K4', 'M4'];
    
    months.forEach((month, index) => {
        const format = index === 3 ? `${fiscalYear}/${month}～ (M${index})` : `${fiscalYear}/${month} (M${index})`;
        worksheet.getCell(cells[index]).value = format;
    });

    // Sample data
    const sampleData: ProductData[] = [
        { id: 1, name: '商品A', price: 1000, threeYearAgoSales: 85, prevPrevYearSales: 90, prevYearSales: 95, sales: [100, 150, 120, 180] },
        { id: 2, name: '商品B', price: 2000, threeYearAgoSales: 65, prevPrevYearSales: 70, prevYearSales: 75, sales: [80, 90, 100, 95] },
        { id: 3, name: '商品C', price: 1500, threeYearAgoSales: 185, prevPrevYearSales: 190, prevYearSales: 195, sales: [200, 180, 220, 210] },
        { id: 4, name: '商品D', price: 3000, threeYearAgoSales: 35, prevPrevYearSales: 40, prevYearSales: 45, sales: [50, 60, 45, 70] },
        { id: 5, name: '商品E', price: 2500, threeYearAgoSales: 140, prevPrevYearSales: 145, prevYearSales: 148, sales: [150, 140, 160, 155] }
    ];

    // Clear the data start row values but keep styles
    worksheet.getRow(DATA_START_ROW).eachCell({ includeEmpty: true }, (cell) => {
        cell.value = null;
    });

    // Set data rows
    sampleData.forEach((item, index) => {
        const row = index + DATA_START_ROW;
        
        // Copy row settings
        if (index !== 0) {
            duplicateRowWithStyles(worksheet, DATA_START_ROW, row);
        }

        // Set basic values
        worksheet.getCell(`A${row}`).value = item.id;
        worksheet.getCell(`B${row}`).value = item.name;
        worksheet.getCell(`C${row}`).value = item.price;
        worksheet.getCell(`D${row}`).value = item.threeYearAgoSales;
        worksheet.getCell(`E${row}`).value = item.prevPrevYearSales;
        worksheet.getCell(`F${row}`).value = item.prevYearSales;

        // Set sales data and calculate percentages
        const salesColumns = ['G', 'I', 'K', 'M'];
        item.sales.forEach((sale, idx) => {
            const salesCell = salesColumns[idx];
            const percentCell = String.fromCharCode(salesCell.charCodeAt(0) + 1);
            
            worksheet.getCell(`${salesCell}${row}`).value = sale;
            worksheet.getCell(`${percentCell}${row}`).value = Number((sale / item.prevYearSales * 100).toFixed(1)) / 100;
        });
    });

    await workbook.xlsx.writeFile('output.xlsx');
    console.log('Excel file has been modified and saved as output.xlsx');
}

function duplicateRowWithStyles(worksheet: ExcelJS.Worksheet, sourceRowNum: number, targetRowNum: number): void {
    const sourceRow = worksheet.getRow(sourceRowNum);
    const targetRow = worksheet.getRow(targetRowNum);

    sourceRow.eachCell({ includeEmpty: true }, (sourceCell, colNumber) => {
        const targetCell = targetRow.getCell(colNumber);
        
        // スタイルのコピー
        targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
        
        // フォント設定のコピー
        if (sourceCell.font) {
            targetCell.font = JSON.parse(JSON.stringify(sourceCell.font));
        }

        // 罫線設定のコピー
        if (sourceCell.border) {
            targetCell.border = JSON.parse(JSON.stringify(sourceCell.border));
        }

        // 保護状態のコピー
        if (sourceCell.protection) {
            targetCell.protection = {
                locked: sourceCell.protection.locked,
                hidden: sourceCell.protection.hidden
            };
        }
    });

    // 行の高さをコピー
    if (sourceRow.height) {
        targetRow.height = sourceRow.height;
    }
}

modifyExcel().catch(err => console.error('Error:', err));