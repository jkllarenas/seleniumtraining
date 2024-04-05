package testapp1;

import java.io.FileOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class PoiDemo {
	
	public static void main(String[] args) {
//		createWorkbook("employees","records");
//		createWorkbook();
		readWorkbook();
//		readWorkbook("employees2","records");
		appendrow("employees2","records");
	}
	
	public static void appendrow (String wb, String ws, String id, String name, String department) {
	}
	public static void readWorkbook(String workbookname,String worksheetname) {
		try {
			File file = new File(workbookname + ".xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheet(worksheetname);

			//loop over rows in sheet
			Iterator<Row> rowiterator = sheet.rowIterator();
			while (rowiterator.hasNext()) {
			Row row = rowiterator.next();
			
			//loop over column in each row
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				System.out.println(cell.getStringCellValue() + "\t");
		} 
			System.out.println("");
		}
		System.out.println("--------end---------");
		} catch (Exception e) {
			e.printStackTrace();
		}
}	
	public static void readWorkbook() {
		try {
			File file = new File ("employees.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheet("employees");
//			XSSFSheet sheet = workbook.getSheetAt(0); kapag hindi mo alam worksheet name

			//loop over rows in sheet
			Iterator<Row> rowiterator = sheet.rowIterator();
			while(rowiterator.hasNext()) {
				Row row = rowiterator.next();
			//loop over columns in each row
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					System.out.print(cell.getStringCellValue() + "\t");
				}
				System.out.println("");
		}
		System.out.println("---------------end-----------");
		} catch (Exception e) {
			e.printStackTrace();
		}
		}
		public static void createWorkbook() {
		//write to xlsx
		//create new instance of workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Employees");
		workbook.createSheet("EmployeesSheet1");
		workbook.createSheet("EmployeesShet2");
		
		//data *no need gawin ito kung file lang ang gagawin
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] {"id","name","department"});
		data.put("2", new Object[] {"1","joseph","mis"});
		data.put("3", new Object[] {"2","ryan","hr"});
		data.put("4", new Object[] {"3","didi","accounting"});
		
		Set<String> keyset = data.keySet();
		
		int rownum = 0;
		//loop
		for(String key:keyset) {
			Row row = sheet.createRow(rownum+=1);
			Object[] obj = data.get(key);
			int cellnum = 0;
			//loop each column in each row
		for(Object o:obj) {
			Cell cell = row.createCell(cellnum+=1);
			cell.setCellValue(o.toString());
			}//end of column loop
		}
		//write file in filesystem
		try {
//			File file = new File ("file.xlsx")
			FileOutputStream out = new FileOutputStream(new File("employees.xlsx")); //ganito ang ginawa ngayon unlike nung nasa upper line na need maglagay ng path dahil sa error na dapat may admin access to create the file
			workbook.write(out);
			out.close();
			System.out.println("write xlsx ok");
		} catch (Exception e) {
			System.out.println(e);
	}
	}
		
	public static void createworkbook(String workbookname,String worksheetname) {
		//write to xlsx
		//create new instance of workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet(worksheetname);
		
		//data *no need gawin ito kung file lang ang gagawin
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] {"id","name","department"});
		data.put("2", new Object[] {"1","joseph","mis"});
		data.put("3", new Object[] {"2","ryan","hr"});
		data.put("4", new Object[] {"3","didi","accounting"});
		
		Set<String> keyset = data.keySet();
		
		int rownum = 0;
		//loop
		for(String key:keyset) {
			Row row = sheet.createRow(rownum+=1);
			Object[] obj = data.get(key);
			int cellnum = 0;
			//loop each column in each row
		for(Object o:obj) {
			Cell cell = row.createCell(cellnum+=1);
			cell.setCellValue(o.toString());
			}//end of column loop
		}
		//write file in filesystem
		try {
//			File file = new File ("file.xlsx")
			FileOutputStream out = new FileOutputStream(new File(workbookname + ".xlsx")); //ganito ang ginawa ngayon unlike nung nasa upper line na need maglagay ng path dahil sa error na dapat may admin access to create the file
			workbook.write(out);
			out.close();
			System.out.println("write xlsx ok");
		} catch (Exception e) {
			System.out.println(e);
	}
	}
		
	}
	
