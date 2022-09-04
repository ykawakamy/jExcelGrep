package excelgrep.core.data;

import java.nio.file.Path;
import java.util.Objects;
import org.apache.poi.ss.util.CellAddress;

public class ExcelPosition {
    public static final int POSITION_NONE = -1;

    Path filePath;

    String sheetName;
    
    String excelType = null;
    int row = POSITION_NONE;
    int col = POSITION_NONE;
    Object extraData = null;

    public enum ExcelPositionType{
        Cell("cell"),
        Shape("[Shape]"),;

        String typeName;
        ExcelPositionType(String typeName) {
            this.typeName = typeName;
        }

        String getValue() {
            return typeName;
        }
    }
    
    public ExcelPosition(Path filePath, String sheetName, int row, int col) {
        this.filePath = filePath;
        this.sheetName = sheetName;
        this.row = row;
        this.col = col;
    }

    public ExcelPosition() {
    }

    public ExcelPosition(Path filePath, String sheetName, String string) {
        this.filePath = filePath;
        this.sheetName = sheetName;
        this.excelType = string;
    }

    public ExcelPosition(Path filePath, String sheetName, ExcelPositionType type) {
        this.filePath = filePath;
        this.sheetName = sheetName;
        this.excelType = type.getValue();
    }
    
    public String getCellPosition() {
        if( this.excelType != null ) {
            return this.excelType;
        }
        return new CellAddress(row, col).formatAsString();
    }
    
    @Override
    public int hashCode() {
        return Objects.hash(col, excelType, extraData, filePath, row, getSheetName());
    }
    @Override
    public boolean equals(Object obj) {
        if (this == obj)
            return true;
        if (obj == null)
            return false;
        if (getClass() != obj.getClass())
            return false;
        ExcelPosition other = (ExcelPosition) obj;
        return col == other.col && Objects.equals(excelType, other.excelType) && Objects.equals(extraData, other.extraData) && Objects.equals(filePath, other.filePath) && row == other.row && Objects.equals(getSheetName(), other.getSheetName());
    }

    public String getSheetName() {
        return sheetName;
    }

    public Path getFilePath() {
        return filePath;
    }

    
    
}
