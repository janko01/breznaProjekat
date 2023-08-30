import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

// I use this function to declare all the methods that I am using from Apache POI.
// These methods help me read files from Excel, write into Excel file and also delete a sheet
public class XLUtility {
    public FileInputStream fi;
    public FileOutputStream fo;
    public XSSFWorkbook workbook;
    public XSSFSheet sheet;
    public XSSFRow row;
    public XSSFCell cell;
    public CellStyle style;
    String path=null;

    //    Whenever I create object of this class, this constructor takes path of the excel file
    XLUtility(String path) {
        this.path = path;
    }
    /*
        This method will return data from the cell
        Cell contains different type of data: formula, String,date, number. BUT we will read
        everything as a string format. I will use special class DataFormatter and formatCellValue method
        that will return the formatted value of a cell as a String regardless of the type. If the cell is empty
        than we are catching that exception and we will catch it and assing empty value to the data variable
     */
    public String getCellData(String sheetName, int rowNum, int colNum) throws IOException {
        fi=new FileInputStream(path);
        workbook=new XSSFWorkbook(fi);
        sheet=workbook.getSheet(sheetName);
        row=sheet.getRow(rowNum);
        cell=row.getCell(colNum);
        DataFormatter formatter = new DataFormatter();
        String data;
        try{
            data = formatter.formatCellValue(cell);
        }
        catch(Exception e){
            data="";
        }
        workbook.close();
        fi.close();
        return data;
    }
    //    With this method, I will write something to the Excel
    public void setCellData(String sheetName, int rowNum, int colNum, String data) throws IOException {
        File xlfile = new File(path);
        if(!xlfile.exists()) { //If file doesn't exist, then I need to create new file
            workbook=new XSSFWorkbook();
            fo=new FileOutputStream(path);
            workbook.write(fo);
        }
//        I am opening file in the input mode
        fi=new FileInputStream(path);
        workbook=new XSSFWorkbook(fi);

        if(workbook.getSheetIndex(sheetName)==-1) //If sheet doesn't exist we need to create it
            workbook.createSheet(sheetName);
        sheet=workbook.getSheet(sheetName);

        if(sheet.getRow(rowNum)==null) //If row doesn't exist then we need to create a row
            sheet.createRow(rowNum);
        row=sheet.getRow(rowNum);

        cell=row.createCell(colNum);
        cell.setCellValue(data);

        fo=new FileOutputStream(path);
        workbook.write(fo);
        workbook.close();
        fi.close();
        fo.close();
    }
    public void deleteSheet() throws FileNotFoundException {
        File xlfile = new File(path);
        xlfile.delete();
    }
}
