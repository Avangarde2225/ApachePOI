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

public class Task1Test {
    @DataProvider(name = "representative")
    public Object[][] representativeData() throws IOException {
        FileInputStream excelFile = new FileInputStream( new File("src/test/resources/representative.xlsx") );
        Workbook wb = new XSSFWorkbook( excelFile );
        Sheet sh = wb.getSheet( "data" );
        Assert.assertNotEquals( sh, null );
        int rowCount = sh.getLastRowNum() - sh.getFirstRowNum();
        System.out.println(rowCount);

        Row firstRow = sh.getRow( 0 );
        Assert.assertNotEquals( firstRow, null );

        int columnCount = firstRow.getLastCellNum() - firstRow.getFirstCellNum();
        System.out.println(columnCount);

        Object[][] resultData = new Object[rowCount][columnCount];

        for(int i = 0; i < rowCount; i++) {
            Row currentRow = sh.getRow( i );
            for(int j = 0; j < columnCount; j++) {
                Cell cell = currentRow.getCell( j );
                resultData[i][j] = cell.toString();
            }
        }

        return  resultData;
    }

    @Test(dataProvider = "representative")
    public void test(String c1, String c2, String c3, String c4){
        System.out.print(c1 + c2 + c3 + c4);
    }
}
