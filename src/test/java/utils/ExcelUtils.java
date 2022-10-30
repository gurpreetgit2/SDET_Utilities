package utils;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class ExcelUtils {

    public static final String pathOfExcel = ".\\data\\Testdata.xlsx";
    XSSFWorkbook workbook;
    XSSFSheet sheet;
    //XSSFRow row;
    //XSSFCell cell;
    DataFormatter formatter;

    /*
    1. constructor to initialize the XSSFWorkbook and Sheet
    2. It is always a good idea to refer sheet with name as the index of sheet are not easy to identify
    3. If you delete all the sheets in EXCEL but the new sheet starts from previous index value plus 1
     */
    public ExcelUtils(String excelPath, String sheetName) throws IOException {
        workbook = new XSSFWorkbook(excelPath);
        sheet = workbook.getSheet(sheetName);
        formatter = new DataFormatter();
    }

    //read the username from column 1 - i.e. cell with index 0 in every row
    public String getUsername(int rowNum){

        return null;
    }
    //read the Password from column 2 - i.e. cell with index 1 in every row
    public String getPassword(int rowNum){

        return null;
    }
    //write the result in Column 3 - i.e. cell with index 2 in every row
    public void setResult(int rowNum){

    }

    //get the records count in a sheet
    public int getRowCount(){
        return sheet.getPhysicalNumberOfRows();
    }

    /*
    1. Method to get the cell value - the most important method
    2. We can use Apache POI DataFormatter 's formatCellValue(Cell cell) method
       as it returns the formatted value of a cell as a String regardless of the cell type.
     */
    public String getCellValue(int rowNum, int cellNum){
        return formatter.formatCellValue(sheet.getRow(rowNum).getCell(cellNum));
    }

    public void setValueForCell(int rowNum, int cellNum, double dataInCell) throws FileNotFoundException {
        //Missing Cell policy to check and create a cell if it does not exist already
        sheet.getRow(rowNum).getCell(cellNum, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(dataInCell);
        try{
            FileOutputStream out=new FileOutputStream("./data/Testdata.xlsx");
            workbook.write(out);
            out.close();
        } catch(Exception e){
            System.out.println("unable to write to excel");
        }
    }

    public static void main(String[] args) throws IOException {
        ExcelUtils obj = new ExcelUtils("./data/Testdata.xlsx","Credentials");
        System.out.println(obj.getRowCount());
        System.out.println(obj.getCellValue(1,0));
        obj.setValueForCell(1,3,200);
    }
}
