package excelServices;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelComparisionServiceMainClass {


	static int file1_column_NO = 0; // Enter first file column value to compare
	static int file2_column_NO = 1; // Enter second file column value to compare
	static int file2_column_NO_additional = 1;
	static int file1_column_NO_additional = 1;
	static String file_name1 = "dbStats_10.34.57.232-22020.xlsx";
	static String file_name2 = "dbStats_10.34.57.232-22020_1.xlsx";
	static String combination_2 = null;

	public static void main(String[] args) {

		try {
			HashMap<Integer, ArrayList> map = new HashMap<Integer, ArrayList>();
			FileInputStream file1 = new FileInputStream(file_name1);
			XSSFWorkbook book1 = new XSSFWorkbook(file1);
			XSSFSheet sheet1 = book1.getSheetAt(0);
			FileInputStream file2 = new FileInputStream(file_name2);
			XSSFWorkbook book2 = new XSSFWorkbook(file2);
			XSSFSheet sheet2 = book2.getSheetAt(0);
			CellType cellType = null;
			int shee1_rowCount = sheet1.getLastRowNum();
			int shee2_rowCount = sheet2.getLastRowNum();
			int row_n = 0;
			System.out.println("First file row count: " + shee1_rowCount + " Second file row count: " + shee2_rowCount);
			Iterator<Row> row_itr = sheet1.rowIterator();
			Row header = row_itr.next(); // Header of the file
			while (row_itr.hasNext()) {
				ArrayList<Object> arr = new ArrayList();
				Row row = row_itr.next();
				row_n = row.getRowNum();
				for (int c = 0; c < row.getLastCellNum(); c++) {
					Cell cell_val = row.getCell(c);
					if (cell_val == null) {
						cellType = CellType.BLANK;
					} else {
						cellType = cell_val.getCellType();
					}
					switch (cellType) {
					case STRING:
						arr.add(cell_val.getStringCellValue());
						break;
					case NUMERIC:
						arr.add(cell_val.getNumericCellValue());
						break;
					case BLANK:
						arr.add("");
						break;
					default:
						break;
					}
				}
				map.put(row_n, arr);
			}

//			System.out.println("Final output " + map);
			excelCompareAndWrite(map, header);
			file1.close();
			file2.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static void excelCompareAndWrite(Map map, Row header1) throws IOException {
		FileInputStream file = new FileInputStream(file_name2);
		Workbook wb = WorkbookFactory.create(file);
		XSSFSheet s = (XSSFSheet) wb.getSheetAt(0);
		Iterator<Row> row_sheet2 = s.iterator();
		Row header_sheet2 = row_sheet2.next();
		int sheet2_cellNo = header_sheet2.getLastCellNum();
		Iterator<Cell> iterator = header1.iterator();
		XSSFRow row1 = s.getRow(0);
		int j = sheet2_cellNo;

		while (iterator.hasNext()) { // Logic to insert the header into the sheet
			Cell cell1 = iterator.next();
			XSSFCell c = row1.createCell(j);
			c.setCellValue(cell1.getStringCellValue());
			j++;
		}
		System.out.println("Headers imported successfully");

		while (row_sheet2.hasNext()) {
			int max_cell = sheet2_cellNo;
			Row sheet2_row = row_sheet2.next();
			Cell cell_value_tobe_compared = sheet2_row.getCell(file1_column_NO);
			
			System.out.println("sheet2: Data is going to be inserted on column no: " + max_cell
					+ ", By taking the data for the comparison : " + cell_value_tobe_compared);
			Iterator<Map.Entry<Integer, ArrayList>> new_Iterator = map.entrySet().iterator();
			while (new_Iterator.hasNext()) {
				Map.Entry<Integer, ArrayList> new_Map = (Map.Entry<Integer, ArrayList>) new_Iterator.next();

				if (cell_value_tobe_compared.getCellType() == CellType.NUMERIC) {
					
					int value1 = ((Double) new_Map.getValue().get(file2_column_NO)).intValue();
					int value2 = (int) cell_value_tobe_compared.getNumericCellValue();
					if (value1 == value2) {
						insertDataMethod(new_Map, max_cell, sheet2_row);
					}
				}
				if (cell_value_tobe_compared.getCellType() == CellType.STRING) {
					String data1 = (String) new_Map.getValue().get(file2_column_NO);
					String data2 = cell_value_tobe_compared.getStringCellValue();
					if(file1_column_NO_additional !=0 )
					{
						Cell data1_additionalString = sheet2_row.getCell(file1_column_NO_additional);
						combination_2 = data2+"|"+data1_additionalString;
					}
					if(file2_column_NO_additional !=0 )
					{
						String data2_additionalString = (String) new_Map.getValue().get(file2_column_NO_additional);
						String combination_1 = data2+"|"+data2_additionalString;
						if (combination_1.equals(combination_2)) {
							insertDataMethod(new_Map, max_cell, sheet2_row);
						}
					}
					else {
					if (data1.equals(data2)) {
						insertDataMethod(new_Map, max_cell, sheet2_row);
					}
					}
				}
				else {
					{
						System.out.println("Cell data type is not   with the expected format !!");
					}
				}

			}
		}
		FileOutputStream out = new FileOutputStream(file_name2);
		wb.write(out);
		out.close();
		System.out.println("Operation completed !!");
	}

	private static void insertDataMethod(Entry<Integer, ArrayList> new_Map, int max_cell, Row sheet2_row) {
		for (Object obj : new_Map.getValue()) {
			Cell cell = sheet2_row.createCell(max_cell);
			if (obj instanceof Double) {
				cell.setCellValue((Double) obj);
			} else if (obj instanceof Integer) {
				cell.setCellValue((Integer) obj);
			} else if (obj instanceof String) {
				cell.setCellValue((String) obj);
			} else {
				cell.setCellValue("");
			}
			max_cell++;
		}
	}
}
