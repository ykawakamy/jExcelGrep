package excelgrep.core;

import java.nio.file.Path;
import java.util.HashSet;
import java.util.Set;
import java.util.regex.Pattern;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.record.cont.ContinuableRecord;
import excelgrep.core.data.ExcelData;
import excelgrep.core.data.ExcelGrepResult;
import excelgrep.core.data.ExcelPosition;
import excelgrep.core.data.ExcelPosition.ExcelPositionType;

public class ExcelGrepHSSFListener implements HSSFListener {
    static Logger log = LogManager.getLogger(ExcelGrepHSSFListener.class);

    ExcelGrepResult result = new ExcelGrepResult();

    private Path filePath;
    private Pattern regex;

    private String currentSheetName;
    private ExcelPosition cellPostion;

    private SSTRecord SST;

    private Set<String> trace_kindRecords = new HashSet<>();

    public ExcelGrepHSSFListener(Path filePath, Pattern regex) {
        super();
        this.filePath = filePath;
        this.regex = regex;
    }

    @Override
    public void processRecord(Record record) {
        if( log.isTraceEnabled() ) {
            log.trace("trace: {} - {}", record.getClass().getSimpleName(), record);
            trace_kindRecords.add(record.getClass().getSimpleName());
        }
        switch (record.getSid()) {
            case BoundSheetRecord.sid:
                onRecord((BoundSheetRecord) record);
                break;
            case BOFRecord.sid:
                onRecord((BOFRecord) record);
                break;
            case EOFRecord.sid:
                onRecord((EOFRecord) record);
                break;
            case ObjRecord.sid:
                onRecord((ObjRecord) record);
                break;
            case BlankRecord.sid:
                onRecord((BlankRecord) record);
                break;
            case DrawingRecord.sid:
                onRecord((DrawingRecord) record);
                break;
            case ExtendedFormatRecord.sid:
            case MergeCellsRecord.sid:
            case StyleRecord.sid:
            case FontRecord.sid:
            case FormatRecord.sid:
            case PaletteRecord.sid:
                // Skip
                break;
            default:
                // TODO need fix
                if (record instanceof CellValueRecordInterface) {
                    onRecord((CellValueRecordInterface) record);
                }
                if (record instanceof CellRecord) {
                    onRecord((CellRecord) record);
                } else if (record instanceof ContinuableRecord) {
                    onRecord((ContinuableRecord) record);
                } else {
                    if(!log.isTraceEnabled()) {
                        log.debug("unhandled: {} - {}", record.getClass().getSimpleName(), record);
                    }
                }
                break;
        }

        

    }

    private void onRecord(CellValueRecordInterface record) {
        cellPostion = new ExcelPosition(filePath, currentSheetName, record.getRow(), record.getColumn());
    }

    private void onRecord(BlankRecord record) {
        cellPostion = new ExcelPosition(filePath, currentSheetName, record.getRow(), record.getColumn());
}

    private void onRecord(DrawingRecord record) {
    }

    private void onRecord(ObjRecord record) {
        for(SubRecord subRecord : record.getSubRecords() ) {
            if( log.isTraceEnabled() ) {
                log.trace("trace sub: {} - {}", subRecord.getClass().getSimpleName(), subRecord);
                trace_kindRecords.add(subRecord.getClass().getSimpleName());
            }            
            if( subRecord instanceof CommonObjectDataSubRecord) {
            }
            if( subRecord instanceof CommonObjectDataSubRecord) {
            }
        }
    }

    private void onRecord(ContinuableRecord record) {
        switch (record.getSid()) {
            case SSTRecord.sid:
                onRecord((SSTRecord) record);
                break;
            case StringRecord.sid:
                onRecord((StringRecord) record);
                break;
            case TextObjectRecord.sid:
                onRecord((TextObjectRecord) record);
                break;
            default:
                log.debug("unhandled: {} - {}", record.getClass().getSimpleName(), record);
        }
    }

    private void onRecord(TextObjectRecord record) {
        cellPostion = new ExcelPosition(filePath, currentSheetName, ExcelPositionType.Shape);
        addResult(record.getStr().toString());
    }

    private void onRecord(StringRecord record) {
        String string = record.getString();

        addResult(string);
    }

    private void onRecord(SSTRecord record) {
        SST = record;
    }

    protected void onRecord(CellRecord record) {

        cellPostion = new ExcelPosition(filePath, currentSheetName, record.getRow(), record.getColumn());
        switch (record.getSid()) {
            case BoolErrRecord.sid:
                onRecord((BoolErrRecord) record);
                break;
            case FormulaRecord.sid:
                onRecord((FormulaRecord) record);
                break;
            case LabelSSTRecord.sid:
                onRecord((LabelSSTRecord) record);
                break;
            case NumberRecord.sid:
                onRecord((NumberRecord) record);
                break;
            case RKRecord.sid:
                onRecord((RKRecord) record);
                break;
        }
    }

    private void onRecord(EOFRecord record) {
        // TODO Auto-generated method stub

    }

    private void onRecord(RKRecord record) {
        // TODO Auto-generated method stub

    }

    private void onRecord(NumberRecord record) {
        // TODO Auto-generated method stub

    }

    private void onRecord(LabelSSTRecord record) {
        String string = SST.getString(record.getSSTIndex()).getString();
        addResult(string);
    }

    private void addResult(String string) {
        if( !regex.matcher(string).find() ) {
            return;
        }
        
        ExcelData data = new ExcelData(cellPostion, string);
        result.add(data);
    }

    private void onRecord(FormulaRecord record) {
        // TODO Auto-generated method stub

    }

    private void onRecord(BoolErrRecord record) {
        // TODO Auto-generated method stub

    }

    private void onRecord(BOFRecord record) {
        switch (record.getType()) {
            case BOFRecord.TYPE_WORKSHEET:
                break;
        }
    }


    private void onRecord(BoundSheetRecord record) {
        currentSheetName = record.getSheetname();

    }

    public void _debug() {
        log.trace("kindRecords: {}", trace_kindRecords);
    }

}
