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
    ExcelGrepResultTableModel model;

    public ExcelGrepSearchWorker(ExcelGrepGuiMain excelGrepGuiMain, String path, String regex) {
        super();
        this.excelGrepGuiMain = excelGrepGuiMain;
        this.path = Paths.get(path);
        this.regex = Pattern.compile(regex);
        this.table = excelGrepGuiMain.table;
        this.model = new ExcelGrepResultTableModel(this.path);

        this.addPropertyChangeListener(new PropertyChangeListener() {
            public void propertyChange(PropertyChangeEvent evt) {
                if ("progress".equals(evt.getPropertyName())) {
                    excelGrepGuiMain.progressBar.setValue((Integer) evt.getNewValue());
                }
            }
        });
    }


    @Override
    protected Void doInBackground() throws Exception {
        table.setEnabled(false);

        String text = "検索中";
        this.excelGrepGuiMain.updateStatusBar(text);
        starttime = System.currentTimeMillis();

        traverseExcelGrep();

        table.setModel(model);

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
        for (ExcelGrepResult it : chunks) {
            for (ExcelData data : it.getResult()) {
                model.addRow(data);
            }
        }

    }


    @Override
    protected void done() {
        long endtime = System.currentTimeMillis();

        this.excelGrepGuiMain.updateStatusBar((endtime - starttime) + "ms");

        table.setEnabled(true);
        table.invalidate();

    }



}
