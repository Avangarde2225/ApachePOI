import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class TestMockData {
    private static final String FILE_NAME = "src/test/resources/MOCK_DATA.xlsx";

    @DataProvider(name = "excelData")
    public Object[][] excelData() throws IOException {
        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet datatypeSheet = workbook.getSheetAt(0);

        int lastRowNum = datatypeSheet.getLastRowNum();
        int firstRowNum = datatypeSheet.getFirstRowNum();
        int rowCount = lastRowNum - firstRowNum;
        int cellCount = datatypeSheet.getRow(1).getLastCellNum() - datatypeSheet.getRow(1).getFirstCellNum();

        Object[][] resultData = new Object[rowCount][cellCount];

        for (int i = 0; i < rowCount; i++) {
            Row row = datatypeSheet.getRow(i);
            if(row != null) {
                for (int j = 0; j < cellCount; j++) {
                    if(row.getCell(j) != null)
                        resultData[i][j] =row.getCell(j).toString();
                }
            }
        }
        return  resultData;
    }

    @Test(dataProvider = "excelData")
    public void test3(String c1, String c2, String c3, String c4, String c5, String c6){
        System.out.println(c1 + c2 + c3 + c4 + c5 + c6);
    }
}
