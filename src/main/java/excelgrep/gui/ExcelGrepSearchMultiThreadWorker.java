package excelgrep.gui;

import java.io.IOException;
import java.nio.file.Files;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;
import javax.swing.JTable;
import excelgrep.core.ExcelGrep;

public class ExcelGrepSearchMultiThreadWorker extends ExcelGrepSearchWorker {

    public ExcelGrepSearchMultiThreadWorker(ExcelGrepGuiMain excelGrepGuiMain, String path, String regex) {
        super(excelGrepGuiMain, path, regex);
    }

    protected void traverseExcelGrep() throws IOException {
        ExecutorService executor = Executors.newWorkStealingPool();
        
        Future<Object> submit = executor.submit( ()->{
            AtomicInteger fileCount = new AtomicInteger(0);
            AtomicInteger completeCount = new AtomicInteger(0);
            Files.walk(path).forEach((it) -> {
                if (isCancelled()) {
                    return;
                }
                fileCount.incrementAndGet();

                executor.submit( ()->{
                    ExcelGrep grep = new ExcelGrep();
                    grep.grepFile(it, regex);
                    publish(grep.getResultSet());
                    
                    completeCount.incrementAndGet();
                    
                    setProgress(100 * completeCount.get() / fileCount.get());
                });
            });
            return null;
        });
        
        try {
            submit.get();
            executor.shutdown();
            while( !executor.awaitTermination(50, TimeUnit.MILLISECONDS)) {
                Thread.yield();
            }
            
        } catch (InterruptedException e) {
        } catch (ExecutionException e) {
            e.printStackTrace();
        }
    }
}
