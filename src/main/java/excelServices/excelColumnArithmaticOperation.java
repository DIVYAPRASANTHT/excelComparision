package excelServices;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelColumnArithmaticOperation {
	static String file = null;

	public void columnDifferenceFinder(String file_name2, String headerName) throws IOException {
		ArrayList list = new ArrayList();
		System.out.println("Arithmatic operation start for excel columns");
		FileInputStream file = new FileInputStream(file_name2);
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		System.out.println(sheet.getSheetName());
		Iterator<Row> itr = sheet.rowIterator();
		String value = "";
		String unit = "";
		double size = 0;
		Row headeRow = itr.next();
		int cell_no = headeRow.getLastCellNum();
		System.out.println("Final value of arithmatic operation will be written in this column no " + cell_no);
		for (Cell cell : headeRow) {
			String value1 = cell.getStringCellValue().trim();
			if (value1.equals(headerName)) {
				list.add(cell.getColumnIndex());
			}
		}
		Cell cell_header = headeRow.createCell(cell_no);
		cell_header.setCellValue("SIZE DIFFERENCE");
		System.out.println("Arithmatic operation would be performed for these 2 columns " + list);
		while (itr.hasNext()) {
			Row row_valueRow = itr.next();
			double sum = 0;
			for (Object number : list) {
				if (row_valueRow.getCell((Integer) number) != null) {
					value = row_valueRow.getCell((Integer) number).getStringCellValue();
					unit = value.replaceAll("[^A-Za-z]", "");
					size = Double.parseDouble(value.replaceAll("[^0-9.9]", ""));
					if (unit.equals("KB")) {
						size = size * 1024;
					} else if (unit.equals("MB")) {
						size = size * 1048576;
					} else if (unit.equals("GB")) {
						size = size * 1073741824;
						;
					}
				} else {
					unit = "KB";
					size = 0;
				}
				sum = size - sum;
			}
			Cell new_cell = row_valueRow.createCell(cell_no);
			String finalsize = "";
			Double finalSizeDouble = sum / 1024;
			DecimalFormat dec_val = new DecimalFormat("0.00");
			finalsize = dec_val.format(finalSizeDouble) + " KB";
			if (finalSizeDouble <= -1024 && finalSizeDouble > -1048576
					|| finalSizeDouble >= 1024 && finalSizeDouble < 1048576) {
				finalSizeDouble = finalSizeDouble / 1024;
				dec_val = new DecimalFormat("0.00");
				finalsize = dec_val.format(finalSizeDouble) + " MB";
			} else if (finalSizeDouble <= -1048576 || finalSizeDouble >= 1048576) {
				finalSizeDouble = finalSizeDouble / 1048576;
				dec_val = new DecimalFormat("0.00");
				finalsize = dec_val.format(finalSizeDouble) + " GB";
			}
			new_cell.setCellValue(finalsize);
		}
		FileOutputStream out = new FileOutputStream(file_name2);
		workbook.write(out);
		out.close();
		System.out.println("Arithmatic operation end for excel columns");
	};
}
