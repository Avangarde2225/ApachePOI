import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class StudentsTest {
    @Test(dataProvider = "studentsData")
    public void test(String c1,String c2,String c3,String c4,String c5,String c6){
        System.out.println(c1  + c2 + c3 + c4 + c5 + c6);
    }

    @DataProvider(name = "studentsData")
    public Object[][] studentsData() throws IOException {
        // Input stream from your excel file
        FileInputStream excelFile = new FileInputStream( new File("src/test/resources/students.xlsx") );
        // using this input stream I create a Workbook object
        Workbook wb = new XSSFWorkbook( excelFile );
        // get the sheet called "data"
        Sheet sheet = wb.getSheet( "data" );
        // check that sheet "data" exists
        Assert.assertNotEquals( sheet, null );
        // find out the row count
        int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
        // based on row count and six column create a two dimensional array
        Object[][] resultData = new Object[rowCount][6];
        // iterate of rows
        for(int i = 0; i < rowCount; i++) {
            // get the current row
            Row currentRow = sheet.getRow( i );
            for(int j = 0; j < 6; j++) {
                // get cell value from current row
                Cell cell = currentRow.getCell( j );
                // assign values to your two dimensional array
                resultData[i][j] = cell.toString();
            }
        }

        return  resultData;

    }
}
