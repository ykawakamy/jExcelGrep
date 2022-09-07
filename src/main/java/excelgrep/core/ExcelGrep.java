package excelgrep.core;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.regex.Pattern;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import excelgrep.core.data.ExcelData;
import excelgrep.core.data.ExcelGrepResult;

public class ExcelGrep {
    static Logger log = LogManager.getLogger(ExcelGrep.class);
    
    ExcelGrepResult result = new ExcelGrepResult();

    public void grepFiles(Path path, Pattern regex) throws IOException {
        Files.walk(path).forEach((it) -> {
            grepFile(it, regex);
        });
    }

    public void grepFile(Path file, Pattern regex) {
        if (file.toFile().isDirectory()) {
            return;
        }
        String filename = file.toString();
        if (filename.endsWith(".xls")) {
            grepHSSF(file, regex);
        } else if (filename.endsWith(".xlsx") || filename.endsWith(".xlsm")) {
            grepXSSF(file, regex);
        }
    }

    private void grepXSSF(Path file, Pattern regex) {
        try ( 
                OPCPackage pkg = OPCPackage.open(file.toString(), PackageAccess.READ);
            ){
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg, false);
            XSSFReader xssfReader = new XSSFReader(pkg);
            StylesTable styles = xssfReader.getStylesTable();

            XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            while (sheets.hasNext()) {
                InputStream stream = sheets.next();
                String sheetName = sheets.getSheetName();
                final XMLReader reader = XMLHelper.newXMLReader();
                ExcelGrepXSSFListener listener = new ExcelGrepXSSFListener(file, regex, sheetName);
                reader.setContentHandler(new XSSFSheetXMLHandler(styles, null, strings, listener, new DataFormatter(), false));
                reader.parse(new InputSource(stream));

                for (XSSFShape it : sheets.getShapes()) {
                    listener.grepShape(it);
                }

                getResultSet().addAll(listener.result);
            }
            
        } catch (Exception e) {
            log.error("unexpected exception." , e);
        }
        
    }

    private void grepHSSF(Path file, Pattern regex) {
        try (POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file.toFile())); InputStream din = fs.createDocumentInputStream("Workbook");) {
            HSSFRequest req = new HSSFRequest();
            ExcelGrepHSSFListener listener = new ExcelGrepHSSFListener(file, regex);
            req.addListenerForAllRecords(listener);
            HSSFEventFactory factory = new HSSFEventFactory();
            factory.processEvents(req, din);

            getResultSet().addAll(listener.result);
            listener._debug();
        } catch (IOException e) {
            log.error("unexpected exception." , e);
        }
    }

    public List<ExcelData> getResult() {
        return result.getResult();
    }

    public ExcelGrepResult getResultSet() {
        return result;
    }
}
