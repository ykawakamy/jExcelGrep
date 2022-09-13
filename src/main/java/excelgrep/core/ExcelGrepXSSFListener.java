package excelgrep.core;

import java.nio.file.Path;
import java.util.HashSet;
import java.util.Set;
import java.util.regex.Pattern;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFAnchor;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import excelgrep.core.data.ExcelData;
import excelgrep.core.data.ExcelGrepResult;
import excelgrep.core.data.ExcelPosition;
import excelgrep.core.data.ExcelPosition.ExcelPositionType;

public class ExcelGrepXSSFListener  extends AbstractResultCollectListener implements SheetContentsHandler {
    static Logger log = LogManager.getLogger(ExcelGrepXSSFListener.class);

    private String sheetName;
    private ExcelPosition cellPostion;

    public ExcelGrepXSSFListener(Path filePath, Pattern regex, String sheetName) {
        super(filePath, regex);

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
        addResult(cellPostion, formattedValue);
    }
    
    public void grepShape(XSSFShape it) {
        if (it instanceof XSSFSimpleShape) {
            XSSFSimpleShape shape = (XSSFSimpleShape)it;
            
            XSSFAnchor anchor = shape.getAnchor();
            if( anchor instanceof XSSFClientAnchor) {
                XSSFClientAnchor clientAncher = (XSSFClientAnchor) anchor;
                short col1 = clientAncher.getCol1();
                int row1 = clientAncher.getRow1();
                cellPostion = new ExcelPosition(filePath, sheetName, row1, col1);

            }else {
                cellPostion = new ExcelPosition(filePath, sheetName, ExcelPositionType.Shape);
            }
            
            addResult(cellPostion, shape.getText());
        }
    }
}
