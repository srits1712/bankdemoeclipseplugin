package datadrivenwritetest;


	
	
	
	
	
	import java.io.FileInputStream;
	import java.io.FileOutputStream;
	import java.io.IOException;
import org.apache.poi.ss.usermodel.DataFormatter;
	import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	public class XLUtils {
		
		public static FileInputStream fi;
		public static FileOutputStream fo;
		public static XSSFWorkbook wb;
		public static XSSFSheet ws;
		public static XSSFRow row;
		public static XSSFCell cell;

		
		/*
		 * Set the File path, open Excel file
		 * @params - Excel Path and Sheet Name
		 */
		public static void setExcelFile(String Path, String SheetName)
				throws Exception {
			try {
				// Open the Excel file
				FileInputStream ExcelFile = new FileInputStream(Path);

				// Access the excel data sheet
				wb = new XSSFWorkbook(ExcelFile);
				ws = wb.getSheet(SheetName);
			} catch (Exception e) {
				throw (e);
			}
		}
		
		public static void createExcelFile(String Path, String SheetName)
				throws Exception {
			try {
				// Open the Excel file
				FileOutputStream fo=new FileOutputStream(Path);
				

				// Access the excel data sheet
				wb = new XSSFWorkbook();
				ws = wb.createSheet(SheetName);
			} catch (Exception e) {
				throw (e);
			}
		}

		public static int getRowCount(String xlfile,String xlsheet) throws IOException 
		{
			fi=new FileInputStream(xlfile);
			wb=new XSSFWorkbook(fi);
			ws=wb.getSheet(xlsheet);
			int rowcount=ws.getLastRowNum();
			wb.close();
			fi.close();
			return rowcount;		
		}
		
		
		public static int getCellCount(String xlfile,String xlsheet,int rownum) throws IOException
		{
			fi=new FileInputStream(xlfile);
			wb=new XSSFWorkbook(fi);
			ws=wb.getSheet(xlsheet);
			row=ws.getRow(rownum);
			int cellcount=row.getLastCellNum();
			wb.close();
			fi.close();
			return cellcount;
		}
		
		public static String getCellData(String xlfile,String xlsheet,int rownum,int colnum) throws IOException
		{
			fi=new FileInputStream(xlfile);
			wb=new XSSFWorkbook(fi);
			ws=wb.getSheet(xlsheet);
			row=ws.getRow(rownum);
			cell=row.getCell(colnum);
			String data;
			try 
			{
				DataFormatter formatter = new DataFormatter();
	            String cellData = formatter.formatCellValue(cell);
	            return cellData;
			}
			catch (Exception e) 
			{
				data="";
			}
			wb.close();
			fi.close();
			return data;
		}
		
		public static void setCellData(String xlfile,String xlsheet,int rownum,int colnum,String data) throws IOException
		{
			fi=new FileInputStream(xlfile);
			wb=new XSSFWorkbook(fi);
			ws=wb.getSheet(xlsheet);
			row=ws.getRow(rownum);
			cell=row.createCell(colnum);
			cell.setCellValue(data);
			fo=new FileOutputStream(xlfile);
			wb.write(fo);		
			wb.close();
			fi.close();
			fo.close();
		}
		
		
		/* * Write in the Excel cell, String Result
		 * @params - Row num and Col num
		 
		public static void setCellData( int RowNum, int ColNum,String Path,String Result)
				throws Exception {
			try {
				row = ws.getRow(RowNum);
				// This should handle null pointer exception if Row does not exist
				if (row == null) {
					row = ws.createRow(RowNum);
				}
				cell = row.getCell(ColNum);
				if (cell == null) {
					cell = row.createCell(ColNum);
					cell.setCellValue(Result);
				} else {
					cell.setCellValue(Result);
				}

				// Open the file to write the results
				FileOutputStream fileOut = new FileOutputStream(Path);

				wb.write(fileOut);
				fo.flush();
				fo.close();
			} catch (Exception e) {
				throw (e);
			}
		}*/
		// returns true if column is created successfully
		public static  boolean addColumn(String xlfile,String xlsheet,String colName){
			//System.out.println("**************addColumn*********************");
			
			try{				
				fi = new FileInputStream(xlfile); 
				wb = new XSSFWorkbook(fi);
				int index = wb.getSheetIndex(xlsheet);
				if(index==-1)
					return false;
				
			/*XSSFCellStyle style = wb.createCellStyle();
			style.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);
			style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);*/
			
			ws=wb.getSheetAt(index);
			
			row = ws.getRow(0);
			if (row == null)
				row = ws.createRow(0);
			
			//cell = row.getCell();	
			//if (cell == null)
			//System.out.println(row.getLastCellNum());
			if(row.getLastCellNum() == -1)
				cell = row.createCell(0);
			else
				cell = row.createCell(row.getLastCellNum());
		        
		        cell.setCellValue(colName);
		     //   cell.setCellStyle(style);
		        
		        fo=new FileOutputStream(xlfile);
				wb.write(fo);
			    fo.close();		    

			}catch(Exception e){
				e.printStackTrace();
				return false;
			}
			
			
			return true;
			
		}
		
		
		public static boolean addSheet(String xlfile,String  sheetname){		
			
			;
			try {
				 wb.createSheet(sheetname);	
				 
				fo = new FileOutputStream(xlfile);
				 wb.write(fo);
			     fo.close();		    
			} catch (Exception e) {			
				e.printStackTrace();
				return false;
			}
			return true;
		}
		//from naveen automaton labs
		
		
		
		
		// returns true if sheet is created successfully else false
			/*	public boolean addSheet(String  sheetname){		
					
					FileOutputStream fileOut;
					try {
						 workbook.createSheet(sheetname);	
						 fileOut = new FileOutputStream(path);
						 workbook.write(fileOut);
					     fileOut.close();		    
					} catch (Exception e) {			
						e.printStackTrace();
						return false;
					}
					return true;
				}
				
				// returns true if sheet is removed successfully else false if sheet does not exist
				public boolean removeSheet(String sheetName){		
					int index = workbook.getSheetIndex(sheetName);
					if(index==-1)
						return false;
					
					FileOutputStream fileOut;
					try {
						workbook.removeSheetAt(index);
						fileOut = new FileOutputStream(path);
						workbook.write(fileOut);
					    fileOut.close();		    
					} catch (Exception e) {			
						e.printStackTrace();
						return false;
					}
					return true;
				}
				// returns true if column is created successfully
				public boolean addColumn(String sheetName,String colName){
					//System.out.println("**************addColumn*********************");
					
					try{				
						fi = new FileInputStream(path); 
						workbook = new XSSFWorkbook(fis);
						int index = workbook.getSheetIndex(sheetName);
						if(index==-1)
							return false;
						
					XSSFCellStyle style = workbook.createCellStyle();
					style.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);
					style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
					
					sheet=workbook.getSheetAt(index);
					
					row = sheet.getRow(0);
					if (row == null)
						row = sheet.createRow(0);
					
					//cell = row.getCell();	
					//if (cell == null)
					//System.out.println(row.getLastCellNum());
					if(row.getLastCellNum() == -1)
						cell = row.createCell(0);
					else
						cell = row.createCell(row.getLastCellNum());
				        
				        cell.setCellValue(colName);
				        cell.setCellStyle(style);
				        
				        fileOut = new FileOutputStream(path);
						workbook.write(fileOut);
					    fileOut.close();		    

					}catch(Exception e){
						e.printStackTrace();
						return false;
					}
					
					return true;
					
					
				}
				// removes a column and all the contents
				public boolean removeColumn(String sheetName, int colNum) {
					try{
					if(!isSheetExist(sheetName))
						return false;
					fis = new FileInputStream(path); 
					workbook = new XSSFWorkbook(fis);
					sheet=workbook.getSheet(sheetName);
					XSSFCellStyle style = workbook.createCellStyle();
					style.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);
					XSSFCreationHelper createHelper = workbook.getCreationHelper();
					style.setFillPattern(HSSFCellStyle.NO_FILL);
					
				    
				
					for(int i =0;i<getRowCount(sheetName);i++){
						row=sheet.getRow(i);	
						if(row!=null){
							cell=row.getCell(colNum);
							if(cell!=null){
								cell.setCellStyle(style);
								row.removeCell(cell);
							}
						}
					}
					fileOut = new FileOutputStream(path);
					workbook.write(fileOut);
				    fileOut.close();
					}
					catch(Exception e){
						e.printStackTrace();
						return false;
					}
					return true;
					
				}
			  // find whether sheets exists	
				public boolean isSheetExist(String sheetName){
					int index = workbook.getSheetIndex(sheetName);
					if(index==-1){
						index=workbook.getSheetIndex(sheetName.toUpperCase());
							if(index==-1)
								return false;
							else
								return true;
					}
					else
						return true;
				}
				
				// returns number of columns in a sheet	
				public int getColumnCount(String sheetName){
					// check if sheet exists
					if(!isSheetExist(sheetName))
					 return -1;
					
					sheet = workbook.getSheet(sheetName);
					row = sheet.getRow(0);
					
					if(row==null)
						return -1;
					
					return row.getLastCellNum();
					
					
					
				}
				//String sheetName, String testCaseName,String keyword ,String URL,String message
				public boolean addHyperLink(String sheetName,String screenShotColName,String testCaseName,int index,String url,String message){
					//System.out.println("ADDING addHyperLink******************");
					
					url=url.replace('\\', '/');
					if(!isSheetExist(sheetName))
						 return false;
					
				    sheet = workbook.getSheet(sheetName);
				    
				    for(int i=2;i<=getRowCount(sheetName);i++){
				    	if(getCellData(sheetName, 0, i).equalsIgnoreCase(testCaseName)){
				    		//System.out.println("**caught "+(i+index));
				    		setCellData(sheetName, screenShotColName, i+index, message,url);
				    		break;
				    	}
				    }


					return true; 
				}
				public int getCellRowNum(String sheetName,String colName,String cellValue){
					
					for(int i=2;i<=getRowCount(sheetName);i++){
				    	if(getCellData(sheetName,colName , i).equalsIgnoreCase(cellValue)){
				    		return i;
				    	}
				    }
					return -1;
					
				}
					
				// to run this on stand alone
				public static void main(String arg[]) throws IOException{
					
					//System.out.println(filename);
					Xls_Reader datatable = null;
					

						 datatable = new Xls_Reader(System.getProperty("user.dir")+"\\src\\com\\qtpselenium\\xls\\Controller.xlsx");
							for(int col=0 ;col< datatable.getColumnCount("TC5"); col++){
								System.out.println(datatable.getCellData("TC5", col, 1)); 
							} 
				}
		}


		*/
	}


