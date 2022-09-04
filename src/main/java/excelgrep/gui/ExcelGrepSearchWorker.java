package excelgrep.gui;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.regex.Pattern;
import javax.swing.SwingWorker;
import excelgrep.core.ExcelGrep;
import excelgrep.core.data.ExcelData;
import excelgrep.core.data.ExcelGrepResult;

public class ExcelGrepSearchWorker extends SwingWorker<Void, ExcelGrepResult> {
    Path path;
    Pattern regex;


    public ExcelGrepSearchWorker(String path, String regex) {
        super();
        this.path = Paths.get(path);
        this.regex = Pattern.compile(regex);
    }


    @Override
    protected Void doInBackground() throws Exception {
        while (!isCancelled()) {
        }
        
        Files.walk(path).forEach((it) -> {
            ExcelGrep grep = new ExcelGrep();
            grep.grepFile(it, regex);
            publish(grep.getResultSet());
        });
        return null;
    }


    @Override
    protected void process(List<ExcelGrepResult> chunks) {
        
    }


    @Override
    protected void done() {
        // TODO Auto-generated method stub
        super.done();
    }
    
    

}
