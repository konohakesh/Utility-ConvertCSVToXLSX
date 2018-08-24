package com.uzumaki.utilities.main;

import java.awt.Color;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.concurrent.TimeUnit;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.uzumaki.utilities.util.BaseCoreUtil;
import com.uzumaki.utilities.util.CommandLineParserUtil;

/**
 * This utility converts Comma Separated Files (.csv) into a Worksheet(.xlsx).
 * If it receives multiple csv files then it converts each CSV file into a sheet
 * inside a single xlsx file in the same order as passed. The name of the sheet
 * should be specified in the name of the CSV file, e.g.
 * MyNovBudget_Summary.xlsx tells that Sheet name should be 'Summary'.
 * (Delimiter are _ and .xlsx)
 */
public class ConvertCsvToWorksheetUtilityMain {

	public static final String WorksheetExt = ".xlsx";

	public static final String CSVFilesPathArg = "-csv";

	public static final String OutputXLSXFilePathArg = "-xlsx";

	public static final String DeleteCSVArg = "-deleteCSV";

	private List<String> csvPaths = null;

	// Map for maintaining CSV sheet name and its corresponding path (maintain
	// insertion order)
	private Map<String, String> csvSheetNamePathMap = new LinkedHashMap<String, String>();

	private String xlsxPath = null;

	private final int Success = 0;

	private final int Failure = 1;

	public static void main(String[] args) {
		ConvertCsvToWorksheetUtilityMain convertCsvToWorksheetUtilityMain = new ConvertCsvToWorksheetUtilityMain();
		convertCsvToWorksheetUtilityMain.convert(args);
	}

	/**
	 * This converts the .csv file into .xlsx file.
	 */
	public void convert(String[] args) {
		int exitStatus = Success;
		long startTime = System.currentTimeMillis();

		if (CommandLineParserUtil.isHelpOption(args)) {
			printHelp();
			System.exit(Success);
		}
		try {
			System.out.println("Validating the arguments...");
			validateArgs(args);

			System.out.println("\nCreating an empty worksheet at '" + xlsxPath + "'...");
			createMetricSpreadSheet(xlsxPath);

			for (Entry<String, String> entry : csvSheetNamePathMap.entrySet()) {
				String sheetName = entry.getKey();
				String csvName = entry.getValue();
				System.out.println("\nCreating '" + sheetName + "' sheet for '" + csvName + "'...");
				convertCSVToXLSX(sheetName, csvName);
			}

			boolean isDeleteCSV = CommandLineParserUtil.isOptionExists(args, DeleteCSVArg);
			if (isDeleteCSV) {
				System.out.println("\nDeleting the CSV file(s)...");

				for (Entry<String, String> entry : csvSheetNamePathMap.entrySet()) {
					BaseCoreUtil.deleteDir(new File(entry.getValue()));
				}
			}

			System.out.println("\nWorksheet is ready at '" + xlsxPath + "'");
		} catch (Exception e) {
			e.printStackTrace();
			exitStatus = Failure;
		} finally {
			wrapUp(startTime, exitStatus);
		}
		System.exit(exitStatus);
	}

	/**
	 * Reads the CSV file and creates a sheet with data in rows, corresponding to it
	 * inside .xlsx file
	 * 
	 * @param sheetName
	 * @param csvFilePath
	 * @throws IOException
	 */
	private void convertCSVToXLSX(String sheetName, String csvFilePath) throws Exception {
		File xlsxFile = new File(xlsxPath);
		BufferedReader csvReader = new BufferedReader(new FileReader(new File(csvFilePath)));
		FileInputStream finXLSX = new FileInputStream(xlsxFile);

		XSSFWorkbook workSheet = new XSSFWorkbook(finXLSX);
		XSSFSheet sheet = workSheet.createSheet(sheetName);

		XSSFCellStyle cellStyleBold = workSheet.createCellStyle();
		XSSFFont font = workSheet.createFont();
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		cellStyleBold.setFont(font);

		int rowNum = 0;
		String line = csvReader.readLine();
		XSSFRow row = null;
		while (line != null) {
			String[] rowCells = StringUtils.splitPreserveAllTokens(line, ",");

			if (rowNum == 0 && rowCells.length >= 2 && !rowCells[1].isEmpty()) {
				cellStyleBold.setFillForegroundColor((new XSSFColor(new Color(145, 191, 238))));
				cellStyleBold.setFillPattern(CellStyle.SOLID_FOREGROUND);
				cellStyleBold.setFillPattern(CellStyle.SOLID_FOREGROUND);

				sheet.setAutoFilter(CellRangeAddress.valueOf("A1:" + ((char) ('A' + (rowCells.length) - 1)) + "1"));
			}

			// Creating a row
			row = sheet.createRow(rowNum);

			int cellNum = 0;
			String oldCell = "";
			for (String cell : rowCells) {
				// Creating a cell
				if ((oldCell == null || oldCell.isEmpty()) && cell.startsWith("\"") && cell.endsWith("\"")) {
					// cell is "124"
					cell = cell.replaceAll("\"", "");
				}
				// These special checks are there to take care of data which itself has comma
				// inside it
				// But to be inside csv they are inside double quotes for e.g. if scanned token
				// is "abc,def,g"
				else if (cell.startsWith("\"") && !cell.endsWith("\"")) {
					// cell is "abc
					cell = cell.replaceAll("\"", "");
					oldCell = cell;
					continue;
				} else if (oldCell != null && (!cell.startsWith("\"") && cell.endsWith("\"") || cell.equals("\""))) {
					// cell is g" or just "
					cell = oldCell + "," + cell;

					cell = cell.replaceAll("\"", "");
					oldCell = null;
				} else if (oldCell != null && !oldCell.isEmpty()) {
					// cell is def
					cell = oldCell + "," + cell;
					oldCell = cell;
					continue;
				}

				createCellInSpreadSheet(row, cellNum, cell, cellStyleBold);
				cellNum++;
			}
			line = csvReader.readLine();

			if (rowNum == 0) {
				XSSFCellStyle cellStyleNormal = workSheet.createCellStyle();
				cellStyleNormal.setAlignment(CellStyle.ALIGN_LEFT);
				cellStyleBold = cellStyleNormal;
			}
			rowNum++;
		}

		if (row != null) {
			int lastColnNo = row.getLastCellNum();
			for (int i = 0; i < lastColnNo; i++) {
				sheet.autoSizeColumn(i);
				try {
					// increasing the width of column as filter widget width eats onto the column
					// names
					sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 600);
				} catch (IllegalArgumentException e) {
					// no need to write into the log for this exception
					// It will come here if excel does not allow width increase of some column if
					// its already at its max post autosize.
					// So we will just don't touch width of this column.
					// java.lang.IllegalArgumentException: The maximum column width for an
					// individual cell is 255 characters.
				}
			}
		}

		FileOutputStream foutXLSX = new FileOutputStream(xlsxFile);
		workSheet.write(foutXLSX);
		csvReader.close();
		finXLSX.close();
		foutXLSX.close();
	}

	/**
	 * Creates a cell in the spreadsheet at cell index in the passed row with
	 * cellValue(String) and cellStyle is applied (optional)
	 */
	public static Cell createCellInSpreadSheet(XSSFRow row, int i, String cellValue, CellStyle cellStyle)
			throws Exception {
		Cell cell = row.createCell(i);

		cellValue = cellValue.trim();

		// capturing numeric values including postive, negative and decimals
		if (cellValue.matches("^-?\\d+(\\.\\d+)?$")) {
			cell.setCellValue(Double.valueOf(cellValue));
		} else {
			cell.setCellValue(cellValue);
		}
		cell.setCellStyle(cellStyle);
		return cell;
	}

	/**
	 * Generating the empty spread sheet now at the start so that each write
	 * operation does not need to create it, it will just open it and add stuff
	 */
	private void createMetricSpreadSheet(String spreadSheetFilePath) throws Exception {
		// Creating the spread sheet
		FileOutputStream out = new FileOutputStream(new File(spreadSheetFilePath));
		XSSFWorkbook wb = new XSSFWorkbook();
		wb.write(out);
		out.close();
	}

	private void wrapUp(long startTime, int exitStatus) {
		if (exitStatus == Success) {
			System.out.println("\nUtility convert_csv_to_xlsx ended SUCCESSFULLY.");
		} else {
			System.out.println("\nUtility convert_csv_to_xlsx FAILED.");
		}

		long endTime = System.currentTimeMillis();
		System.out.println(
				"Duration: " + Long.toString(TimeUnit.MILLISECONDS.toSeconds(endTime - startTime)) + " second(s).");
	}

	/**
	 * Validates the arguments passed to the utility
	 */
	public void validateArgs(String[] args) throws Exception {

		List<String> invalidInputArgs = CommandLineParserUtil.getInvalidArguments(getSupportedArguments(), args);
		if (invalidInputArgs != null) {
			throw new Exception("\nThese arguments are not supported by this utility:"
					+ Arrays.toString(invalidInputArgs.toArray()));
		}

		csvPaths = CommandLineParserUtil.buildList(args, CSVFilesPathArg);
		if (csvPaths == null || csvPaths.size() <= 0) {
			throw new Exception("\nPlease specify atleast one CSV file path in -csv argument.");
		} else {
			int i = 0;
			for (String path : csvPaths) {
				if (!path.endsWith(".csv")) {
					throw new Exception("\nThis CSV file path should end with .csv extension: " + path);
				} else {
					File csvFile = new File(path);
					if (!csvFile.exists()) {
						throw new Exception("This CSV file does not exist: " + path);
					}

					int k = path.lastIndexOf("_");
					if (k != -1) {
						String sheetName = path.substring(k + 1, path.length() - 4);
						csvSheetNamePathMap.put(sheetName, path);
					} else {
						csvSheetNamePathMap.put("Sheet" + i, path);
						i++;
					}
				}
			}
		}

		xlsxPath = CommandLineParserUtil.getValueForKey(args, OutputXLSXFilePathArg);
		if (xlsxPath == null || xlsxPath.isEmpty()) {
			throw new Exception("\nPlease specify the path of .xlsx file in -output arguments.");
		} else if (!xlsxPath.endsWith(WorksheetExt)) {
			throw new Exception("\nXLSX file path should end with " + WorksheetExt + " extension: " + xlsxPath);
		}
	}

	public String[] getSupportedArguments() {
		String[] mSuppportedArguments = { "-h", CSVFilesPathArg, OutputXLSXFilePathArg, DeleteCSVArg };
		return mSuppportedArguments;
	}

	/**
	 * Prints help.
	 */
	private static void printHelp() {
		StringBuffer buffer = new StringBuffer();
		buffer.append("Usage: convert_csv_to_xlsx\n");
		buffer.append("      -csv=\"<path1 of a csv file>,<path2 of a csv file>\"\n");
		buffer.append("      -xlsx=\"<path for the consolidated xlsx file>\"\n");
		buffer.append("      [-deleteCSV] [-h]\n\n");
		buffer.append("This utility converts Comma Separated Files (.csv) into a Worksheet(.xlsx).\n");
		buffer.append(
				"If it receives multiple CSV files then it converts each CSV file into a sheet inside a single xlsx file in the same order as passed.\n"); //$NON-NLS-1$
		buffer.append(
				"The name of the sheet should be specified in the name of the CSV file, e.g. Metrics_Steps.xlsx tells that sheet name is 'Steps'. (Delimiters are _ and .xlsx)\n\n"); //$NON-NLS-1$
		buffer.append("Arguments:\n");
		buffer.append("  -csv=             Full path to a CSV file. (Comma separated paths if multiple CSV files.)\n");
		buffer.append("  -xlsx=            Full path of the .xlsx file which should get generated.\n");
		buffer.append(
				"                    Full path should mention the name of the .xlsx file, e.g. C:\\MyExcel.xlsx \n");
		buffer.append("  -deleteCSV        Delete the CSV file(s) post XLSX file generation\n");
		buffer.append("  -h                Displays help for this utility.\n");

		System.out.println(buffer.toString());
	}
}
