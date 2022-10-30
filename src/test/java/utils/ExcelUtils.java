package utils;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public class ExcelUtils {

    public static final String pathOfExcel = ".\\data\\Testdata.xlsx";
    XSSFWorkbook workbook;
    XSSFSheet sheet;

    //constructor to initialize the XSSFWorkbook and Sheet
    public ExcelUtils(String excelPath, String sheetName) throws IOException {
        workbook = new XSSFWorkbook(excelPath);
        sheet = workbook.getSheet(sheetName);
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
        System.out.println("Total number of rows are: "+sheet.getPhysicalNumberOfRows());
        System.out.println("Total number of rows are: "+sheet.getLastRowNum());
        return sheet.getPhysicalNumberOfRows();
        //return sheet.getLastRowNum();
    }

    public static void main(String[] args) throws IOException {
        ExcelUtils obj = new ExcelUtils("./data/Testdata.xlsx","Credentials");
        obj.getRowCount();
    }
}
