package datadrivenwritetest;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ExcelWrite {



	public static void main(String[] args) throws Exception {

//writing the comments 

		/*   FileOutputStream file = new FileOutputStream("E:\\bindu network\\Wspaces\\Workspace_Bindu\\dummy\\Excelcreated.xlsx");

		   XSSFWorkbook workbook = new XSSFWorkbook();
		  XSSFSheet sheet = workbook.createSheet("data");

		   for (int i = 0; i <= 5; i++) {
		   XSSFRow row = sheet.createRow(i);
		   for (int j = 0; j < 3; j++) {
		    row.createCell(j).setCellValue("abcdef");
		   }
		  }

		   workbook.write(file);
		  System.out.println("writing excel is completed");

		  */
	String path="E:\\bindu network\\Wspaces\\Workspace_Bindu\\dummy\\LoginData.xlsx";


	   FileInputStream file = new FileInputStream(path);
	  XSSFWorkbook workbook = new XSSFWorkbook(file);
	  XSSFSheet sheet = workbook.getSheet("Sheet1");

	   int rownum = sheet.getLastRowNum(); // returns number of rows present in excel sheet
	  int colcount = sheet.getRow(0).getLastCellNum(); // returns number of cells present in a row

	   for (int i = 0; i < rownum; i++) {
	   XSSFRow row = sheet.getRow(i);

	    for (int j = 0; j < colcount; j++) {
	    String value = row.getCell(j).toString(); // reading the data from cell
	    System.out.print(value + " ");
	   }

	   

	   }
	

	
	String path1="E:\\bindu network\\Wspaces\\Workspace_Bindu\\dummy\\Excelcreated.xlsx";
		FileOutputStream fo = new FileOutputStream(path1);

		   XSSFWorkbook wb = new XSSFWorkbook();
		  XSSFSheet s= workbook.createSheet("data");
		 /* XLUtils.setCellData(path1, "Results",0,0,"username");
			XLUtils.setCellData(path1, "Results",0,1,"password");
			XLUtils.setCellData(path1, "Results",0,2,"Status");
	*/
			
			
		  for (int i = 1; i <=5; i++) {
				   XSSFRow row = s.createRow(i);
				  
				   for (int j = 0; j < 2; j++) {
				  
				    
			
			XLUtils.setCellData(path1, "Result", i,j, value);
				   }
				  }
			 workbook.write(fo);
			  System.out.println("writing excel is completed");

	}
}

		

