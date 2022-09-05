package excelgrep.gui;

import java.util.List;
import java.util.Vector;
import javax.swing.table.AbstractTableModel;
import excelgrep.core.data.ExcelData;
import excelgrep.core.data.ExcelGrepResult;

@SuppressWarnings("serial")
public class ExcelGrepResultTableModel extends AbstractTableModel {
    protected List<String> columnIdentifiers;
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
        ExcelData data = result.getResult(rowIndex);
        switch( columnIndex ) {
            case 0:
                return data.getPosition().getFilePath();
            case 1:
                return data.getPosition().getSheetName();
            case 2:
                return data.getPosition().getCellPosition();
            case 3:
                return data.getValue().getStrings();
        }
        return null;
    }

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

}