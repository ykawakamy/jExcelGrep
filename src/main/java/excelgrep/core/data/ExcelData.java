package excelgrep.core.data;

import java.util.Objects;

public class ExcelData {
    ExcelPosition position = new ExcelPosition();
    ExcelValue value = new ExcelValue();
    
    public ExcelData(ExcelPosition cellPostion, String string) {
        this.position = cellPostion;
        this.getValue().add(string);
    }
    @Override
    public int hashCode() {
        return Objects.hash(getPosition(), getValue());
    }
    @Override
    public boolean equals(Object obj) {
        if (this == obj)
            return true;
        if (obj == null)
            return false;
        if (getClass() != obj.getClass())
            return false;
        ExcelData other = (ExcelData) obj;
        return Objects.equals(getPosition(), other.getPosition()) && Objects.equals(getValue(), other.getValue());
    }
    public ExcelPosition getPosition() {
        return position;
    }
    public ExcelValue getValue() {
        return value;
    }
    
    
}
