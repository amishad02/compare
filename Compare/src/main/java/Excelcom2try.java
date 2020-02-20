import java.awt.Color;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Reporter;


public class Excelcom2try  {
	static FileOutputStream opstr1 = null;
	    static XSSFCellStyle cellStyleRed = null;
	    static   SXSSFWorkbook sxssfWorkbook = null;
	    static   SXSSFSheet sheet1new = null;
	    static   SXSSFRow row3edit = null;
  
    @SuppressWarnings("resource")
    public static void main(String[] args) {
    	System.out.println("open");
    try {
        // Create new file for Result 
        XSSFWorkbook workbook = new XSSFWorkbook();
        FileOutputStream fos = new FileOutputStream(new File("F:\\ResultFile.xlsx"));
        workbook.write(fos);
        workbook.close();
        Thread.sleep(2000);
        // get input for 2 compare excel files
        FileInputStream excellFile1 = new FileInputStream(new File("F:\\UAT_Rel.xlsx"));
        FileInputStream excellFile2 = new FileInputStream(new File("F:\\Prod_Rel.xlsx"));
        // Copy file 2 for result to highlight not equal cell
        FileSystem system = FileSystems.getDefault();
        Path original = system.getPath("F:\\Prod_Rel.xlsx");
        Path target = system.getPath("F:\\ResultFile.xlsx");
        try {
            // Throws an exception if the original file is not found.
            Files.copy(original, target, StandardCopyOption.REPLACE_EXISTING);
            Reporter.log("Successfully Copy File 2 for result to highlight not equal cell");
         
        } catch (IOException ex) {
            Reporter.log("Unable to Copy File 2 ");
         
        }
        Thread.sleep(2000);
        FileInputStream excelledit3 = new FileInputStream(new File("F:\\ResultFile.xlsx"));
        // Create Workbook for 2 compare excel files
        XSSFWorkbook workbook1 = new XSSFWorkbook(excellFile1);
        XSSFWorkbook workbook2 = new XSSFWorkbook(excellFile2);
        // Temp workbook 
        XSSFWorkbook workbook3new = new XSSFWorkbook();
        //XSSF cellStyleRed as  SXSSFWorkbook cannot have cellstyle  color
        cellStyleRed = workbook3new.createCellStyle();
        cellStyleRed.setFillForegroundColor(IndexedColors.RED.getIndex());
        cellStyleRed.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Get first/desired sheet from the workbook to compare both excel sheets
        XSSFSheet sheet1 = workbook1.getSheetAt(0);
        XSSFSheet sheet2 = workbook2.getSheetAt(0);
        //XSSFWorkbook workbook3new temp convert to SXSSFWorkbook
        // keep 100 rows in memory, exceeding rows will be flushed to disk
        sxssfWorkbook = new SXSSFWorkbook(workbook3new, 100);
        sheet1new = (SXSSFSheet) sxssfWorkbook.createSheet();
        // Compare sheets
        if (compareTwoSheets(sheet1, sheet2, sheet1new)) {

        	System.out.println("\\n\\nThe two excel sheets are Equal");
            
        } else {
          
            System.out.println("\\n\\nThe two excel sheets are Not Equal");

        }

        // close files
        excellFile1.close();
        excellFile2.close();
        excelledit3.close();
        opstr1.close();
    } catch (Exception e) {
        e.printStackTrace();
    }
    Reporter.log("Successfully Close All files");
    System.out.println("close");
}

// Compare Two Sheets
public static boolean compareTwoSheets(XSSFSheet sheet1, XSSFSheet sheet2, SXSSFSheet sheet1new) throws IOException {
    int firstRow1 = sheet1.getFirstRowNum();
    int lastRow1 = sheet1.getLastRowNum();
    boolean equalSheets = true;
    for (int i = firstRow1; i <= lastRow1; i++) {

        Reporter.log("\n\nComparing Row " + i);
     
        XSSFRow row1 = sheet1.getRow(i);
        XSSFRow row2 = sheet2.getRow(i);

        row3edit = (SXSSFRow) sheet1new.getRow(i);
        if (!compareTwoRows(row1, row2, row3edit)) {
            equalSheets = false;
            // Write if not equal
// Get error here java.lang.NullPointerException for row3edit.setRowStyle(cellStyleRed);
            //if disable test is completed Successfully without writing result file 
          //  row3edit.setRowStyle(cellStyleRed);
            System.out.println("Row " + i + " - Not Equal");
        
            // break;
        } else {
        	System.out.println("Row " + i + " - Equal");
           
        }
    }
    // Write if not equal 
    opstr1 = new FileOutputStream("F:\\ResultFile.xlsx");
    sxssfWorkbook.write(opstr1);

    opstr1.close();

    return equalSheets;
}

// Compare Two Rows
public static boolean compareTwoRows(XSSFRow row1, XSSFRow row2, SXSSFRow row3edit) throws IOException {
    if ((row1 == null) && (row2 == null)) {
        return true;
    } else if ((row1 == null) || (row2 == null)) {
        return false;
    }

    int firstCell1 = row1.getFirstCellNum();
    int lastCell1 = row1.getLastCellNum();
    boolean equalRows = true;

    // Compare all cells in a row

    for (int i = firstCell1; i <= lastCell1; i++) {
        XSSFCell cell1 = row1.getCell(i);
        XSSFCell cell2 = row2.getCell(i);
        if (!compareTwoCells(cell1, cell2)) {
            equalRows = false;
            System.out.println("       Cell " + i + " - NOt Equal " + cell1 + "  ===  " + cell2);
           
            break;
        } else {
        	System.out.println("       Cell " + i + " - Equal " + cell1 + "  ===  " + cell2);
         
        }
    }
    return equalRows;
}

// Compare Two Cells
@SuppressWarnings("deprecation")
public static boolean compareTwoCells(XSSFCell cell1, XSSFCell cell2) {
    if ((cell1 == null) && (cell2 == null)) {
        return true;
    } else if ((cell1 == null) || (cell2 == null)) {
        return false;
    }

    boolean equalCells = false;
    int type1 = cell1.getCellType();
    int type2 = cell2.getCellType();
    if (type2 == type1) {
        if (cell1.getCellStyle().equals(cell2.getCellStyle())) {
            // Compare cells based on its type
            switch (cell1.getCellType()) {
            case HSSFCell.CELL_TYPE_FORMULA:
                if (cell1.getCellFormula().equals(cell2.getCellFormula())) {
                    equalCells = true;
                } else {
                }
                break;

            case HSSFCell.CELL_TYPE_NUMERIC:
                if (cell1.getNumericCellValue() == cell2.getNumericCellValue()) {
                    equalCells = true;
                } else {
                }
                break;
            case HSSFCell.CELL_TYPE_STRING:
                if (cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
                    equalCells = true;
                } else {
                }
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                if (cell2.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
                    equalCells = true;

                } else {
                }
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                if (cell1.getBooleanCellValue() == cell2.getBooleanCellValue()) {
                    equalCells = true;
                } else {
                }
                break;
            case HSSFCell.CELL_TYPE_ERROR:
                if (cell1.getErrorCellValue() == cell2.getErrorCellValue()) {
                    equalCells = true;
                } else {
                }
                break;
            default:
                if (cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
                    equalCells = true;
                } else {
                }
                break;
            }
        } else {
            return false;
        }
    } else {
        return false;
    }
    return equalCells;
}
}
