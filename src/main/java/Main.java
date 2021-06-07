import EasyXLS.*;
import EasyXLS.Constants.DataType;

import java.awt.*;
import java.time.Instant;

public class Main {
    public static void main(String[] args) {
        try {
            Instant start =Instant.now();
            System.out.println(start + " start");
            System.out.println("Tutorial 05");
            System.out.println("----------");

            ExcelDocument workbook = new ExcelDocument(1);

            System.out.println("Reading file C:\\Samples\\test.xlsb");
            workbook.easy_getOptions().setCalculateFormulas(false);
            // Set the sheet names (2)
            workbook.easy_getSheetAt(0).setSheetName("first");
            workbook.easy_getSheetAt(1).setSheetName("Second tab");

            ExcelTable xlsFirstTable = ((ExcelWorksheet) workbook.easy_getSheetAt(0)).easy_getExcelTable();

            ExcelStyle style = new ExcelStyle("Verdana", 8, true, true, Color.YELLOW);
            for (int column = 0; column < 5; column++) {
                xlsFirstTable.easy_getCell(0, column).setValue("Column " + (column + 1));
            }
            // Add data in cells for report values
            Instant start2 =Instant.now();
            System.out.println(start2 + " start cycle");
            for (int row = 0; row < 50000; row++) {
                for (int column = 0; column < 21; column++) {
                    xlsFirstTable.easy_getCell(row + 1, column).setStyle(style);
                    xlsFirstTable.easy_getCell(row + 1, column).setValue("Data " + (row + 1) + ", " + (column + 1));
                    xlsFirstTable.easy_getCell(row+1,column).setDataType(DataType.STRING);
                }
            }
            Instant start3 =Instant.now();
            System.out.println(start3 + " finis cycle");
            // Export the Excel file
            Instant startWrite = Instant.now();
            System.out.println(startWrite + "   start Xslb");
            workbook.easy_WriteXLSBFile("C:\\Samples\\1\\test.xlsb");
            Instant finishWrite = Instant.now();
            System.out.println(finishWrite + "   finish Xlsb");
            // Confirm export of Excel file
            if (workbook.easy_getError().equals(""))
                System.out.println("File successfully created.");
            else
                System.out.println("Error encountered: " + workbook.easy_getError());
            // Dispose memory
            workbook.Dispose();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
}