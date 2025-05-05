package com.excel.excelSheetDemo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;


public class excelDemo {
    public static void main(String[] args) {
        excelDemo obj = new excelDemo();
        obj.readData();
        obj.writeData();
    }

    public void readData() {
        try {
            FileInputStream file = new FileInputStream("sampleSheet.xlsx");
            Workbook workbook = new XSSFWorkbook(file);

            Sheet sheet = workbook.getSheetAt(0);

            for(Row row : sheet) {
                for(Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING :
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;

                        case BOOLEAN :
                            System.out.print(cell.getBooleanCellValue() + "\t");
                            break;

                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;

                        case FORMULA:
                            System.out.print(cell.getCellFormula() + "\t");
                            break;

                        default:
                            System.out.print("UNKNOWN DATA!");
                    }
                }
                System.out.println();

            }
        }
        catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }
    public void writeData() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Employees");
        //sheet.setColumnWidth(0, 4000);
        //sheet.setColumnWidth(1, 4000);
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);

        Row header = sheet.createRow(0);

        // Create header cell
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        headerStyle.setFont(font);

        Cell headerCell = header.createCell(0);
        headerCell.setCellValue("Name");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(1);
        headerCell.setCellValue("Age");
        headerCell.setCellStyle(headerStyle);

        // Create content cell
        CellStyle contentStyle = workbook.createCellStyle();
        contentStyle.setWrapText(true);

        Row row = sheet.createRow(1);
        Cell contentCell = row.createCell(0);
        contentCell.setCellValue("Vikash L B");
        contentCell.setCellStyle(contentStyle);

        contentCell = row.createCell(1);
        contentCell.setCellValue(22);
        contentCell.setCellStyle(contentStyle);

        // write data into a file
        try {
            File curr = new File("data.xlsx");
            FileOutputStream outputStream = new FileOutputStream(curr);
            workbook.write(outputStream);
            workbook.close();
            System.out.println("File Created - 'data.xlsx'  ");
        }
        catch (Exception e){
            System.out.println(e.getMessage());
        }
    }
}
