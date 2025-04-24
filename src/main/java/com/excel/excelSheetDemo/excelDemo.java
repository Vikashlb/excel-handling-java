package com.excel.excelSheetDemo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;


public class excelDemo {
    public static void main(String[] args) {
        excelDemo obj = new excelDemo();
        obj.readData();
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
}
