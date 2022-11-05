package excelgrep.core.data;

import java.nio.file.Path;
import java.util.Objects;
import excelgrep.core.data.ExcelPosition.ExcelPositionType;


/**
 * 検索結果
 * <p></p>
 */
public class ExcelData {
    ExcelPosition position = new ExcelPosition();
    ExcelValue value = new ExcelValue();

    public static ExcelData newFileNameMatch(Path file) {
        return new ExcelData(new ExcelPosition(file, "", ExcelPositionType.FileNameMatch), "");
    }

    public static ExcelData newLoadFailure(Path file, Exception e) {
        return new ExcelData(new ExcelPosition(file, "", ExcelPositionType.LoadFailure), e.toString());
    }

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
