package excelgrep.gui;

import java.nio.file.Path;
import java.util.List;
import java.util.Vector;
import javax.swing.table.AbstractTableModel;
import excelgrep.core.data.ExcelData;
import excelgrep.core.data.ExcelGrepResult;

@SuppressWarnings("serial")
public class ExcelGrepResultTableModel extends AbstractTableModel {
    protected List<String> columnIdentifiers;
    protected Path basePath;
    
    ExcelGrepResult result = new ExcelGrepResult();

    public ExcelGrepResultTableModel(String[] columnNames) {
        columnIdentifiers = List.of(columnNames);
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
                return basePath.relativize(data.getPosition().getFilePath());
            case 1:
                return data.getPosition().getSheetName();
            case 2:
                return data.getPosition().getCellPosition();
            case 3:
                return data.getValue().getStrings();
        }
        return null;
    }

    public ExcelData getRow(int rowIndex) {
        return result.getResult(rowIndex);
    }

    @Override
    public String getColumnName(int column) {
        Object id = null;
        // This test is to cover the case when
        // getColumnCount has been subclassed by mistake ...
        if (column < columnIdentifiers.size() && (column >= 0)) {
            id = columnIdentifiers.get(column);
        }
        return (id == null) ? super.getColumnName(column)
                            : id.toString();
    }

    public void addRow(ExcelData data) {
        result.add(data);
    }
    
    public void clear(Path path) {
        basePath = path;
        result.clear();
    }

}
