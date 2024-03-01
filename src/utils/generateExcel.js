// Importar las librerías necesarias
import ExcelJS from 'exceljs';
import Company from '../company/company.model.js';

// Función para generar el reporte Excel
export const generateExcelReport = async () => {
    let workbook = new ExcelJS.Workbook();
    let worksheet = workbook.addWorksheet('Companies');

    // Obtener todas las empresas registradas desde la base de datos
    let companies = await Company.find();

    // Definir las columnas del reporte con estilos
    worksheet.columns = [
        { header: 'Name', key: 'name', width: 20 },
        { header: 'Category', key: 'categoryBusiness', width: 20 },
        { header: 'Years Carrer', key: 'yearsCareer', width: 20 },
    ];

    // Eestilos para las celdas del encabezado
    worksheet.getRow(1).eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFA07A' },
        };
    });

    // Agregar los datos de las empresas al reporte con estilos
    companies.forEach((company, index) => {
        const rowIndex = index + 2; // Empezar en la segunda fila después del encabezado
        const row = worksheet.getRow(rowIndex);
        row.values = [company.name, company.categoryBusiness, company.yearsCareer];

        // Establecer estilos para las celdas de datos
        row.eachCell((cell) => {
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        });
    });

    // Guardar el archivo Excel con los estilos aplicados
    await workbook.xlsx.writeFile('Reports/companies_report.xlsx');
};