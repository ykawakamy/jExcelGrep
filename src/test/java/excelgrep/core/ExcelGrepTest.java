package excelgrep.core;

import static org.junit.jupiter.api.Assertions.*;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.regex.Pattern;
import org.junit.jupiter.api.Test;
import excelgrep.core.ExcelGrep;
import excelgrep.core.data.ExcelData;
import excelgrep.core.data.ExcelValue;

class ExcelGrepTest {

    @Test
    void test() throws IOException {
        ExcelGrep testee = new ExcelGrep();
        testee.grepFiles(Paths.get("./src/test/input/case1/"), Pattern.compile("フロー"));
        
        assertTrue( ()->{
            ExcelData act = testee.getResultSet().getResult(0);
            assertEquals("29", act.getPosition().getSheetName());
            assertEquals("C1", act.getPosition().getCellPosition() );
            assertEquals(new ExcelValue("1.02  我が国のエネルギーバランス・フロー概要"), act.getValue());
            return true;
        });
    }

    @Test
    void test2() throws IOException {
        ExcelGrep testee = new ExcelGrep();
        testee.grepFiles(Paths.get("./src/test/input/case1/"), Pattern.compile("ガソリン"));
        
        assertTrue( ()->{
            ExcelData act = testee.getResultSet().getResult(0);
            assertEquals("29", act.getPosition().getSheetName());
            assertEquals("[Shape]", act.getPosition().getCellPosition() );
            assertEquals(new ExcelValue("ガソリン\t\t 1,413"), act.getValue());
            return true;
        });
    }

    @Test
    void test2_1() throws IOException {
        ExcelGrep testee = new ExcelGrep();
        testee.grepFiles(Paths.get("./src/test/input/case2/"), Pattern.compile("フロー"));
        
        assertTrue( ()->{
            ExcelData act = testee.getResultSet().getResult(0);
            assertEquals("29", act.getPosition().getSheetName());
            assertEquals("C1", act.getPosition().getCellPosition() );
            assertEquals(new ExcelValue("1.02  我が国のエネルギーバランス・フロー概要"), act.getValue());
            return true;
        });
    }

    @Test
    void test2_2() throws IOException {
        ExcelGrep testee = new ExcelGrep();
        testee.grepFiles(Paths.get("./src/test/input/case2/"), Pattern.compile("ガソリン"));
        
        assertTrue( ()->{
            ExcelData act = testee.getResultSet().getResult(0);
            assertEquals("29", act.getPosition().getSheetName());
            assertEquals("[Shape]", act.getPosition().getCellPosition() );
            assertEquals(new ExcelValue("ガソリン\t\t 1,413"), act.getValue());
            return true;
        });
    }

    @Test
    void test3() throws IOException {
        ExcelGrep testee = new ExcelGrep();
        testee.grepFiles(Paths.get("./src/test/input/case3/"), Pattern.compile("ガソリン"));
        
        assertTrue( ()->{
            ExcelData act = testee.getResultSet().getResult(0);
            assertEquals("Sheet1", act.getPosition().getSheetName());
            assertEquals("B74", act.getPosition().getCellPosition() );
            assertEquals(new ExcelValue("ガソリン\t\t 1,413"), act.getValue());
            return true;
        });
    }

}
