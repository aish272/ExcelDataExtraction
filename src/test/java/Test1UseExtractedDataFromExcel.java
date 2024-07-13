import utility.ExtractDataFromExcel;

import java.io.IOException;
import java.util.List;

public class Test1UseExtractedDataFromExcel {

    public static void main(String[] args) throws IOException {
        ExtractDataFromExcel extractor = new ExtractDataFromExcel();
        List<String> inputDataList = extractor.getExcelData("Purchase");
        inputDataList.forEach(data ->System.out.println(data));
    }

}
