import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {
	// Test Case :- Identify Test cases column by scanning the entire row
	// Once column is identified then scan entire test cases column to identify
	// Purchase Test Case row
	// after you grab purchase test case row , pull all the data of that row and
	// feed into test.
	
	
	public ArrayList<String> getData(String testcaseName) throws IOException {
		ArrayList<String> a = new ArrayList<String>();

		// Step 1:Create Object for XSSFWorkbook class
		FileInputStream fis = new FileInputStream("D:\\Personal_docs\\datademo.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		System.out.println("-------------");

		// 2.Get Access to the Sheet
		int sheets = workbook.getNumberOfSheets();
		for (int i = 0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				// Identify Testcases column by scanning the entire 1st row.
				Iterator<Row> rows = sheet.iterator(); // sheet is a collection of rows
				Row firstrow = rows.next();
				// Iterate in the first row now
				Iterator<Cell> ce = firstrow.cellIterator();// row is a collection of cells.
				int k = 0;
				int column = 0;
				while (ce.hasNext()) {
					Cell value = ce.next();
					if (value.getStringCellValue().equalsIgnoreCase("Testcases")) {
						// Getting Desired Column index
						column = k;

					}

					k++;
				}

				System.out.println(column);
				// Once column is identified then scan entire test cases column to identify
				// Purchase Test Case row
				while(rows.hasNext()) {
					Row r = rows.next();
					//for every row only get me the value of the cell of column index
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName))
					{
						// after you grab purchase test case row , pull all the data of that row and feed into test.
						
						Iterator<Cell> cv = r.cellIterator();
						while(cv.hasNext()) {
							Cell c = cv.next();
							if(c.getCellTypeEnum()==CellType.STRING) {
								
								a.add(c.getStringCellValue());
							}else {
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
							
						}
					}		
				}

			}

		}
		return a;
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		
	}

}
