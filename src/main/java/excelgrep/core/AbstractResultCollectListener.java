package excelgrep.core;



import java.nio.file.Path;
import java.util.regex.Pattern;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import excelgrep.core.data.ExcelData;
import excelgrep.core.data.ExcelGrepResult;
import excelgrep.core.data.ExcelPosition;

public class AbstractResultCollectListener {

    final protected ExcelGrepResult result = new ExcelGrepResult();
    final protected Path filePath;
    final protected Pattern regex;



    public AbstractResultCollectListener(Path filePath, Pattern regex) {
        super();
        this.filePath = filePath;
        this.regex = regex;
    }



    protected void addResult(ExcelPosition position, String string) {
        if( !regex.matcher(string).find() ) {
            return;
        }
        
        ExcelData data = new ExcelData(position, string);
        result.add(data);
    }

}
