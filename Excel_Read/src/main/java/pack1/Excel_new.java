package pack1;

package pack1;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_new
{
static FileInputStream f;
static XSSFWorkbook w;
static XSSFSheet sh;
public static String getStringData(int a,int b) throws IOException //method to get string data
{
f=new FileInputStream("C:\\Users\\Merin\\Documents\\Test.xlsx");  //reading data from a file in the form of bytes
w=new XSSFWorkbook(f);        //creating workbook instance that refers to .xls file  
sh=w.getSheet("Sheet1");       //to get the first sheet
Row r=sh.getRow(a);//to get the row
Cell c=r.getCell(b);//to get the cell
return c.getStringCellValue();
}
public static String getIntegerData(int a,int b) throws IOException //method to get string data
{
f=new FileInputStream("C:\\Users\\91944\\Documents\\Book1.xlsx");//reading data from a file in the form of bytes
w=new XSSFWorkbook(f); //instance to refer workbook
sh=w.getSheet("Sheet1"); //to get the first sheet
Row r=sh.getRow(a);//to get the row
Cell c=r.getCell(b);//to get the cell
int x=(int)c.getNumericCellValue(); //converting double data to integer
return String.valueOf(x);
}
}



