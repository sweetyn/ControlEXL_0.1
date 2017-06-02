package nvg.mm.re;

import java.io.File;
import java.io.IOException;
//import java.text.SimpleDateFormat;
//import java.util.Calendar;
//import java.util.Date;
//import java.util.Locale;



import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.TimeZone;

import jxl.Cell;
import jxl.DateCell;
//import jxl.DateFormulaCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
//import jxl.write.Label;
//import jxl.write.WritableSheet;
//import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class estimate {
	public static void main(String[] args) throws RowsExceededException, WriteException {
		
		try {
			
			// Read a workbook from a file
			Workbook workbook = Workbook.getWorkbook(new File("WK21.xls"));
			
			// Read a sheet from the workbook
			Sheet sheet = workbook.getSheet(0);
			
			//String[] sheetname = workbook.getSheetNames();
			//System.out.println(sheetname[0]);
			
			// Identify number of rows in the sheet
			int numRows = sheet.getRows();
						
			// Identify number of columns in the sheet
			int numCols = sheet.getColumns();
			//System.out.println(numCols);
			
		
			
			Cell gettempDate = sheet.getCell(1, 0);
			System.out.println("used getContent() " + gettempDate.getContents());
			System.out.println("type is " + gettempDate.getType());
			
			DateCell dateCell = (DateCell) gettempDate;
			System.out.println("getdate() is " + dateCell.getDate() + " Note: This time is off & needs correction");
			
			
			TimeZone gmtZone = TimeZone.getTimeZone("GMT");
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
			sdf.setTimeZone(gmtZone);
			
			System.out.println("result after SDF " + sdf.format(dateCell.getDate()));
			
				
			/*
			// Identify number of WK
			for (int i = 0; i < numCols; i++) {
				Cell cellCols = sheet.getCell(i, 1);
								
				String cellWK = cellCols.getContents();
				//System.out.println(cellWK);
				//int nameWK = cellWK.indexOf("WK");
				//System.out.println(nameWK);
				
				//Identify model for this week(Col)
				if (cellWK.indexOf("WK") == 0){					
					//System.out.println(cellWK);
					
					//check this week
					if (cellWK.indexOf("WK23")==0){									
						//Cell cellTemp = sheet.getCell(i+1, 0);
						
						// Identify model for this week(row)
						for (int j = 0; j < numRows; j++) {
							Cell cellTemp = sheet.getCell(i+1, j);	
							String cellVal = cellTemp.getContents();
							//cellTemp.getCellFormat();
							//cellTemp.equals(null);
							
							if (cellVal == ""){					
							}
							else {
								//Cell cellTemp = sheet.getCell(i+1, j);
								//cell cellTemp = sheet.
								Cell gettempModel = sheet.getCell(2, j);
								String tempModel = gettempModel.getContents();
								
								Cell gettempProject = sheet.getCell(3, j);
								String tempProject = gettempProject.getContents();
								
								Cell gettempDate = sheet.getCell(i+1, 3);
								//Cell gettempDate = sheet.getCell(11, 3);
								
								//String tempDate = gettempDate.get
										
								//DateCell tempDate = (DateCell) gettempDate;
								//Date date = gettempDate.get
								//Date tempDate = gettempDate.getDate();
								//String test = tempDate.toString();
												
								System.out.println(tempModel + " (" +tempProject + ") " + cellVal + " : " + "date" );							//	System.out.println(tempModel + " (" +tempProject + ") " + cellVal + " : " + test );
							//	System.out.println(j);
							}
						}
						
						//String cellid = cellRows.getContents();
					//	System.out.println("i");
						
						
						
						
						
						
						
						
						
						
						
					}
					else{
						//System.out.println("i");
						System.out.println(cellWK.substring(0, 2));
					}
					//System.out.println(cellWK.substring(0, 2));
				}
				else {
					//System.out.println(i);
				}
			}

			
			
			*/
			
			
			// Identify model for this week(row)
						/*for (int i = 0; i < numRows; i++) {
							Cell cellRows = sheet.getCell(0, i);
											
							String cellid = cellRows.getContents();
							System.out.println(cellid);
							
							if (cellid == ""){					
							}
							else {
								System.out.println(i);
							}
						}*/
			
			
						
			//Get Start Date
			//String[] sheetname = workbook.getSheetNames();
			//Sheet sheet = workbook.getSheet(0);
			//Cell c0 = sheet1.getCell(3, 1);
			//System.out.println(c0);
			
			
			/*// Read the first sheet
			String[] sheetname = workbook.getSheetNames();
			Sheet sheet1 = workbook.getSheet(0);
			
			
			// Find running model
			for (int i=0; i==300; i++){
				Cell tempCell = sheet1.getCell(i,0);
				String tempCellID = tempCell.getContents();
				System.out.println(tempCellID);
				if (tempCellID=="0"){
					System.out.println(i);
				}
			}
			*/
			
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
