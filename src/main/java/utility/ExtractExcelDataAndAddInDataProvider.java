package utility;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;

import java.io.FileInputStream;
import java.io.IOException;

public class ExtractExcelDataAndAddInDataProvider {

    /**
     * This will extract excel data and return through data provider
     *
     * @return
     */
    @DataProvider (name="inputDataFromExcel")
    public Object[][] returnExcelDataThroughDataProvider() throws IOException {
        FileInputStream fis = new FileInputStream("/Users/aish/Documents/ExcelDataExtraction/src/main/resources/InputData.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet dataSheet =workbook.getSheet("Data");
        int colCount = dataSheet.getRow(0).getLastCellNum();
        int rowCount = dataSheet.getLastRowNum();
        Object[][] twoDimArray = new Object[rowCount-1][colCount-1]; //Subtracting 1 from rowCount and column count because we are excluding header
        DataFormatter formatter = new DataFormatter();
        for(int i=0;i<rowCount-1;i++)
        {
            XSSFRow targetRow = dataSheet.getRow(i+1);
            for (int j= 0;j<colCount-1;j++)
            {
                twoDimArray[i][j] = formatter.formatCellValue(targetRow.getCell(j+1));
            }
        }

        return twoDimArray;
    }
}
