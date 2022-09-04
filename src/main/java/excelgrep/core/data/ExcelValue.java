package excelgrep.core.data;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class ExcelValue {
    List<String> values = new ArrayList<>();

    
    public ExcelValue(String... values) {
        super();
        this.values = new ArrayList<>(List.of(values));
    }

    
    public void add(String string) {
        this.values.add(string);
    }
    
    @Override
    public int hashCode() {
        return Objects.hash(values);
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj)
            return true;
        if (obj == null)
            return false;
        if (getClass() != obj.getClass())
            return false;
        ExcelValue other = (ExcelValue) obj;
        return Objects.equals(values, other.values);
    }


    @Override
    public String toString() {
        StringBuilder builder = new StringBuilder();
        builder.append("ExcelValue [values=");
        builder.append(values);
        builder.append("]");
        return builder.toString();
    }
    
    public String getStrings() {
        StringBuilder builder = new StringBuilder();
        for( String it : values ) {
            builder.append(it);
        }
        return builder.toString();
    }
}
