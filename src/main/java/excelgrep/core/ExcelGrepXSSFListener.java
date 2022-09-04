package excelgrep.core;

import java.nio.file.Path;
import java.util.HashSet;
import java.util.Set;
import java.util.regex.Pattern;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import excelgrep.core.data.ExcelData;
import excelgrep.core.data.ExcelGrepResult;
import excelgrep.core.data.ExcelPosition;
import excelgrep.core.data.ExcelPosition.ExcelPositionType;

public class ExcelGrepXSSFListener implements SheetContentsHandler {
    static Logger log = LogManager.getLogger(ExcelGrepXSSFListener.class);

    ExcelGrepResult result = new ExcelGrepResult();

    private Path filePath;
    private Pattern regex;

    private String sheetName;
    private ExcelPosition cellPostion;


    private Set<String> trace_kindRecords = new HashSet<>();

    public ExcelGrepXSSFListener(Path filePath, Pattern regex, String sheetName) {
        super();
        this.filePath = filePath;
        this.regex = regex;
        this.sheetName = sheetName;
    }
    

    @Override
    public void startRow(int rowNum) {
        // ...
    }

    @Override
    public void endRow(int rowNum) {
        // ...
    }

    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        cellPostion = new ExcelPosition(filePath, sheetName, cellReference);
        addResult(formattedValue);
    }
    
    private void addResult(String string) {
        if( !regex.matcher(string).find() ) {
            return;
        }
        
        ExcelData data = new ExcelData(cellPostion, string);

        result.add(data);
    }


    public void grepShape(XSSFShape it) {
        if (it instanceof XSSFSimpleShape) {
            XSSFSimpleShape shape = (XSSFSimpleShape)it;
            cellPostion = new ExcelPosition(filePath, sheetName, ExcelPositionType.Shape);
            addResult(shape.getText());
        }
    }
}
