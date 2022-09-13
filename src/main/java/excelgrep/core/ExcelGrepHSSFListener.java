package excelgrep.core;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ddf.EscherRecord;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.record.AbstractEscherHolderRecord;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BlankRecord;
import org.apache.poi.hssf.record.BoolErrRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.CellRecord;
import org.apache.poi.hssf.record.CellValueRecordInterface;
import org.apache.poi.hssf.record.CommonObjectDataSubRecord;
import org.apache.poi.hssf.record.DrawingGroupRecord;
import org.apache.poi.hssf.record.DrawingRecord;
import org.apache.poi.hssf.record.EOFRecord;
import org.apache.poi.hssf.record.ExtendedFormatRecord;
import org.apache.poi.hssf.record.FontRecord;
import org.apache.poi.hssf.record.FormatRecord;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.MergeCellsRecord;
import org.apache.poi.hssf.record.MulBlankRecord;
import org.apache.poi.hssf.record.NoteRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.ObjRecord;
import org.apache.poi.hssf.record.PaletteRecord;
import org.apache.poi.hssf.record.RKRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.RowRecord;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.record.StringRecord;
import org.apache.poi.hssf.record.StyleRecord;
import org.apache.poi.hssf.record.SubRecord;
import org.apache.poi.hssf.record.TextObjectRecord;
import org.apache.poi.hssf.record.cont.ContinuableRecord;
import excelgrep.core.data.ExcelPosition;

public class ExcelGrepHSSFListener extends AbstractResultCollectListener implements HSSFListener {
    static Logger log = LogManager.getLogger(ExcelGrepHSSFListener.class);

    private String currentSheetName;
    private ExcelPosition cellPostion;

    private SSTRecord SST;
    private List<BoundSheetRecord> boundSheetRecords;
    private int activeSheetIdx = 0;

    public ExcelGrepHSSFListener(Path filePath, Pattern regex) {
        super(filePath, regex);
        this.boundSheetRecords = new ArrayList<>();
    }

    @Override
    public void processRecord(Record record) {

        try {
            internalProcessRecord(record);
        } catch (Exception e) {
            log.error("unexpected exception.", e);
        }
    }

    private void internalProcessRecord(Record record) {
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
            case MulBlankRecord.sid:
                onRecord((MulBlankRecord) record);
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
            case RowRecord.sid:
                onRecord((RowRecord) record);
                break;

            case NoteRecord.sid:
                onRecord((NoteRecord) record);
                break;
            default:
                if (record instanceof CellValueRecordInterface) {
                    onRecord((CellValueRecordInterface) record);
                }
                if (record instanceof CellRecord) {
                    onRecord((CellRecord) record);
                } else if (record instanceof ContinuableRecord) {
                    onRecord((ContinuableRecord) record);
                } else if (record instanceof AbstractEscherHolderRecord) {
                    onRecord((AbstractEscherHolderRecord) record);
                } else {
                    log.trace("unhandled: {} - {}", record.getClass().getSimpleName(), record);
                }
                break;
        }
    }

    private void onRecord(NoteRecord record) {
        record.getShapeId();

    }

    private void onRecord(MulBlankRecord record) {
        cellPostion = new ExcelPosition(filePath, currentSheetName, record.getRow(), record.getFirstColumn());

    }

    private void onRecord(RowRecord record) {
        cellPostion = new ExcelPosition(filePath, currentSheetName, record.getRowNumber(), record.getFirstCol());

    }

    private void onRecord(AbstractEscherHolderRecord record) {
        if (record instanceof DrawingGroupRecord) {
            DrawingGroupRecord casted = (DrawingGroupRecord) record;
            casted.processChildRecords();
        }
        List<EscherRecord> escherRecords = record.getEscherRecords();
        onEscherRecord(escherRecords);
    }

    private void onEscherRecord(List<EscherRecord> escherRecords) {
        if (escherRecords.isEmpty()) {
            return;
        }
        for (EscherRecord it : escherRecords) {
            onEscherRecord(it.getChildRecords());
            log.debug("{}", it);
        }
    }

    private void onRecord(CellValueRecordInterface record) {
        cellPostion = new ExcelPosition(filePath, currentSheetName, record.getRow(), record.getColumn());
    }

    private void onRecord(DrawingRecord record) {}

    private void onRecord(ObjRecord record) {
        for (SubRecord subRecord : record.getSubRecords()) {
            if (subRecord instanceof CommonObjectDataSubRecord) {
                CommonObjectDataSubRecord casted = (CommonObjectDataSubRecord) subRecord;
            }
            if (subRecord instanceof CommonObjectDataSubRecord) {
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
        // cellPostion = new ExcelPosition(filePath, currentSheetName, ExcelPositionType.Shape);
        addResult(cellPostion, record.getStr().toString());
    }

    private void onRecord(StringRecord record) {
        String string = record.getString();

        addResult(cellPostion, string);
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
        addResult(cellPostion, string);
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
                currentSheetName = boundSheetRecords.get(activeSheetIdx).getSheetname();
                activeSheetIdx++;
                break;
        }
    }


    private void onRecord(BoundSheetRecord record) {
        boundSheetRecords.add(record);

    }

}
