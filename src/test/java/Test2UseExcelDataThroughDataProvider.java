import org.testng.annotations.Test;
import utility.ExtractExcelDataAndAddInDataProvider;

public class Test2UseExcelDataThroughDataProvider {


    @Test(dataProvider = "inputDataFromExcel", dataProviderClass = ExtractExcelDataAndAddInDataProvider.class)
    public void printData(String input1, String input2, String input3) {
        System.out.println("InputData1" + input1 + " " + "InputData2" + input2 + " " + "InputData3" + input3);
    }

}

