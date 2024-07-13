package utility;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class ExtractDataFromExcel {

    public ArrayList<String> getExcelData(String testcaseName) throws IOException {
        /* Scenario:
          Identify Testcase column and get the column number.
          Scan the whole column and get the row with 'Purchase' string.
          Use the row to extract all the test data.
         */
        FileInputStream fis = new FileInputStream("/Users/aish/Documents/ExcelDataExtraction/src/main/resources/InputData.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        int sheetNum = wb.getNumberOfSheets();
        Sheet dataSheet = null;
        for (int i = 0; i < sheetNum; i++)
            if (wb.getSheetAt(i).getSheetName().equals("Data")) {
                dataSheet = wb.getSheetAt(i);
                break;
            }
        Iterator<Row> rows = dataSheet.iterator(); //sheet is a collection of rows
        Row firstRow = rows.next();
        Iterator<Cell> firstRowsCells = firstRow.iterator();
        int testcasesColIndex = 0;
        while (firstRowsCells.hasNext()) {
            Cell c = firstRowsCells.next();
            if (c.getStringCellValue().equals("Testcases")) {
                testcasesColIndex = c.getColumnIndex();
            }
        }
        Row targetRow;
        ArrayList<String> inputDataList = new ArrayList<>();
        while (rows.hasNext()) {
            targetRow = rows.next();
            if (targetRow.getCell(testcasesColIndex).getStringCellValue().equals(testcaseName)) {
                Iterator<Cell> purchaseDataCells = targetRow.cellIterator(); //Row is a collection of cells
                while (purchaseDataCells.hasNext()) {
                    Cell c = purchaseDataCells.next();
                    CellType type = c.getCellType();
                    switch (type)
                    {
                        case STRING -> inputDataList.add(c.getStringCellValue());
                        case NUMERIC -> inputDataList.add(String.valueOf(c.getNumericCellValue()));
                        case BOOLEAN -> inputDataList.add(String.valueOf(c.getBooleanCellValue()));
                    }

                }
            }
        }
        return inputDataList;


    }
}
