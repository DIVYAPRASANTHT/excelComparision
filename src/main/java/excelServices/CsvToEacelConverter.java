package excelServices;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

import org.apache.commons.lang.NumberUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CsvToEacelConverter {
	public static final char CSV_FILE_DELIMITER = ',';
	public void convertCsvToExcel(String strSource, String strDestination, String extension)
            throws IllegalArgumentException, IOException {
 
        Workbook workBook = null;
        FileOutputStream fos = null;
 
        // Check that the source file exists.
        File sourceFile = new File(strSource);
        if (!sourceFile.exists()) {
            throw new IllegalArgumentException("The source CSV file cannot be found at " + sourceFile);
        }
 
        // Check that the destination folder exists to save the Excel file.
        File destination = new File(strDestination);
        if (!destination.exists()) {
            throw new IllegalArgumentException(
                    "The destination directory " + destination + " for the converted Excel file does not exist.");
        }
        if (!destination.isDirectory()) {
            throw new IllegalArgumentException(
                    "The destination " + destination + " for the Excel file is not a directory/folder.");
        }
 
        // Getting BufferedReader object
        BufferedReader br = new BufferedReader(new FileReader(sourceFile));
 
        // Getting XSSFWorkbook or HSSFWorkbook object based on excel file format
        if (extension.equals(".xlsx")) {
            workBook = new XSSFWorkbook();
        } else {
            workBook = new HSSFWorkbook();
        }
 
        XSSFSheet sheet = (XSSFSheet) workBook.createSheet("Sheet");
 
        String nextLine;
        int rowNum = 0;
        while ((nextLine = br.readLine()) != null) {
            Row currentRow = sheet.createRow(rowNum++);
            String rowData[] = nextLine.split(String.valueOf(CSV_FILE_DELIMITER));
            for (int i = 0; i < rowData.length; i++) {
                if (NumberUtils.isDigits(rowData[i])) {
                    currentRow.createCell(i).setCellValue(Integer.parseInt(rowData[i]));
                } else if (NumberUtils.isNumber(rowData[i])) {
                    currentRow.createCell(i).setCellValue(Double.parseDouble(rowData[i]));
                } else {
                    currentRow.createCell(i).setCellValue(rowData[i]);
                }
            }
        }
        String filename = sourceFile.getName();
        filename = filename.substring(0, filename.lastIndexOf('.'));
        File generatedExcel = new File(strDestination, filename + extension);
        fos = new FileOutputStream(generatedExcel);
        workBook.write(fos);
 
        try {
            // Closing workbook, fos, and br object
            workBook.close();
            fos.close();
            br.close();
 
        } catch (IOException e) {
            System.out.println("Exception While Closing I/O Objects");
            e.printStackTrace();
        }
 
    }
}
