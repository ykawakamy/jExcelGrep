package excelgrep.core.data;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

public class ExcelGrepResult {
    List<ExcelData> result = new ArrayList<>();

    public void addAll(ExcelGrepResult data) {
        getResult().addAll(data.getResult());
    }

    public void add(ExcelData data) {
        getResult().add(data);
        
    }

    public List<ExcelData> getResult() {
        return result;
    }

    public ExcelData getResult(int i) {
        return result.get(i);
    }
}
