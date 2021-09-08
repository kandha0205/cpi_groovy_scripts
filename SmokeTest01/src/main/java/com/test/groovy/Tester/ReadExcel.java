package com.test.groovy.Tester;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

public class ReadExcel {

    public static void main(String[] args ) throws Exception {

        InputStream input = new FileInputStream(new File("C:\\Users\\M000077\\OneDrive - Uniper SE\\CPI\\uniper\\English_Receive_Single_Line.xlsx"));
      //  Workbook workbook = WorkbookFactory.create(new File("C:\\Users\\M000077\\OneDrive - Uniper SE\\CPI\\uniper\\English_Receive_Single_Line.xlsx"));
        Workbook workbook = WorkbookFactory.create(input);

        Sheet sheet = workbook.getSheetAt(0) ;
        Row row = sheet.getRow(6);
         }
}
