package com.automation.tests;

import org.apache.poi.ss.usermodel.*;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;

public class WriteIntoExcelFile {

    @Test
    public void writeIntoFileTest()throws IOException {
        FileInputStream inputStream = new FileInputStream ( "VytrackTestUsers.xlsx" );
        Workbook workbook =WorkbookFactory.create ( inputStream );
        inputStream.close ();

        Sheet sheet = workbook.getSheet ( "QA3-short" );
        Row row = sheet.getRow ( 1 ); // 2nd row
        Cell cell = row.getCell ( 5 );// get result column

        System.out.println ("BEFORE: "+cell.getStringCellValue ());
        cell.setCellValue ( "PASSED: ");
        System.out.println ("AFTER: "+cell.getStringCellValue ());

        Row firstRow = sheet.getRow ( 0 );
        Cell newCell = firstRow.createCell ( 6 );
        newCell.setCellValue ("Date Execution");

        Row secondRow = sheet.getRow ( 1 );
        Cell newCell2 = secondRow.createCell ( 6 );// create new cell
        newCell2.setCellValue ( LocalDateTime.now ().toString () );//I will write date and time info into the cell

        FileOutputStream outputStream = new FileOutputStream ("VytrackTestUsers.xlsx"  ); // close to excell file
        workbook.write ( outputStream );
        workbook.close ();
        outputStream.close ();
    }
}
