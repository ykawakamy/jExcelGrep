package excelgrep.gui;

import java.nio.file.Path;
import java.util.List;
import java.util.Vector;
import javax.swing.table.AbstractTableModel;
import excelgrep.core.data.ExcelData;
import excelgrep.core.data.ExcelGrepResult;

@SuppressWarnings("serial")
public class ExcelGrepResultTableModel extends AbstractTableModel {
    protected Path basePath;
    
    ExcelGrepResult result = new ExcelGrepResult();

    public ExcelGrepResultTableModel(Path path) {
        basePath = path;
    }

    public ExcelGrepResultTableModel() {
    }

    @Override
    public int getRowCount() {
        return result.getResult().size();
    }

    @Override
    public int getColumnCount() {
        return 4;
    }

    @Override
    public Object getValueAt(int rowIndex, int columnIndex) {
        ExcelData data = getRow(rowIndex);
        switch( columnIndex ) {
            case 0:
                return basePath.relativize(data.getPosition().getFilePath()).getParent();
            case 1:
                return basePath.relativize(data.getPosition().getFilePath()).getFileName();
            case 2:
                return data.getPosition().getSheetName();
            case 3:
                return data.getPosition().getCellPosition();
            case 4:
                return data.getValue().getStrings();
        }
        return null;
    }

    public ExcelData getRow(int rowIndex) {
        return result.getResult(rowIndex);
    }

    public void addRow(ExcelData data) {
        result.add(data);
    }
}
