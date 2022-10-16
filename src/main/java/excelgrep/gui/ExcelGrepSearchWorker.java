package excelgrep.gui;

import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.regex.Pattern;
import javax.swing.JTable;
import javax.swing.SwingWorker;
import excelgrep.core.ExcelGrep;
import excelgrep.core.data.ExcelData;
import excelgrep.core.data.ExcelGrepResult;

class ExcelGrepSearchWorker extends SwingWorker<Void, ExcelGrepResult> {
    private final ExcelGrepGuiMain excelGrepGuiMain;
    final Path path;
    final Pattern regex;
    final JTable table;

    long starttime;

    public ExcelGrepSearchWorker(ExcelGrepGuiMain excelGrepGuiMain, String path, String regex, JTable table) {
        super();
        this.excelGrepGuiMain = excelGrepGuiMain;
        this.path = Paths.get(path);
        this.regex = Pattern.compile(regex);
        this.table = table;
        
        this.addPropertyChangeListener(
                new PropertyChangeListener() {
                    public  void propertyChange(PropertyChangeEvent evt) {
                        if ("progress".equals(evt.getPropertyName())) {
                            excelGrepGuiMain.progressBar.setValue((Integer)evt.getNewValue());
                        }
                    }
                });
    }


    @Override
    protected Void doInBackground() throws Exception {
        table.setEnabled(false);
        ExcelGrepResultTableModel model = (ExcelGrepResultTableModel) table.getModel();
        model.clear(path);
        
        String text = "検索中";
        this.excelGrepGuiMain.updateStatusBar(text);
        starttime = System.currentTimeMillis();

        traverseExcelGrep();

        return null;
    }


    protected void traverseExcelGrep() throws IOException {
        Files.walk(path).forEach((it) -> {
            if (isCancelled()) {
                return;
            }

            ExcelGrep grep = new ExcelGrep();
            grep.grepFile(it, regex);
            publish(grep.getResultSet());
        });
    }


    @Override
    protected void process(List<ExcelGrepResult> chunks) {
        ExcelGrepResultTableModel model = (ExcelGrepResultTableModel) table.getModel();
        for (ExcelGrepResult it : chunks) {
            for (ExcelData data : it.getResult()) {
                model.addRow(data);
            }
        }

    }


    @Override
    protected void done() {
        long endtime = System.currentTimeMillis();

        this.excelGrepGuiMain.updateStatusBar((endtime-starttime) + "ms");

        table.setEnabled(true);
        table.invalidate();

    }



}