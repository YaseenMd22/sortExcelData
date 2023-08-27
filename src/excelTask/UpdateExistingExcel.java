package excelTask;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
//import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.ArrayList;

public class UpdateExistingExcel {

	public static Properties prop;
	public static String workingDir = System.getProperty("user.dir");
	public static String propFilePath = workingDir + "//config.properties";

	public static void main(String[] args) {
		
		try {
			prop = new Properties();
			FileInputStream fp = new FileInputStream(propFilePath);
			prop.load(fp);
		} catch (IOException e) {
			e.printStackTrace();
		}

		String inputFile = prop.getProperty("filePath"); // Replace with your input Excel file

		try (FileInputStream fis = new FileInputStream(inputFile); Workbook workbook = new XSSFWorkbook(fis)) {

			Sheet sheet1 = workbook.getSheetAt(0);
			Sheet sheet2 = workbook.getSheetAt(1);

			Sheet sheet3 = workbook.createSheet("Sheet 3");
			int newRowNum = 0;

			Map<String, List<Row>> idToRowsMap = new LinkedHashMap<>(); // Use LinkedHashMap for insertion order

			processSheet(idToRowsMap, sheet1, "H");
			processSheet(idToRowsMap, sheet2, "I");

			for (Map.Entry<String, List<Row>> entry : idToRowsMap.entrySet()) {
				String id = entry.getKey();
				List<Row> idRows = entry.getValue();

				for (Row row : idRows) {
					newRowNum = addRowToSheet3(sheet3, newRowNum, idRows.indexOf(row) == 0 ? "H" : "I", row);
				}
			}

			try (FileOutputStream fos = new FileOutputStream(inputFile)) {
				workbook.write(fos);
			}

			System.out.println("Sheet 3 created and added to the same Excel file.");

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static void processSheet(Map<String, List<Row>> idToRowsMap, Sheet sheet, String sheetType) {
		for (Row row : sheet) {
			if (row.getRowNum() == 0) {
				// Skip header row
				continue;
			}

			String id = getStringValue(row.getCell(0));
			List<Row> idRows = idToRowsMap.getOrDefault(id, new ArrayList<>());
			idRows.add(row);
			idToRowsMap.put(id, idRows);
		}
	}

	private static String getStringValue(Cell cell) {
		if (cell != null) {
			switch (cell.getCellType()) {
			case STRING:
				return cell.getStringCellValue();
			case NUMERIC:
				return String.valueOf(cell.getNumericCellValue());
			default:
				return "";
			}
		}
		return "";
	}

	private static int addRowToSheet3(Sheet sheet, int rowNum, String sheetName, Row sourceRow) {
		Row newRow = sheet.createRow(rowNum++);
		Cell sheetNameCell = newRow.createCell(0);
		Cell idCell = newRow.createCell(1);

		sheetNameCell.setCellValue(sheetName);
		idCell.setCellValue(getStringValue(sourceRow.getCell(0)));

		for (int colIndex = 1; colIndex < sourceRow.getLastCellNum(); colIndex++) {
			Cell dataCell = newRow.createCell(colIndex + 1);
			Cell sourceCell = sourceRow.getCell(colIndex);
			copyCellValue(sourceCell, dataCell);
		}

		return rowNum;
	}

	private static void copyCellValue(Cell sourceCell, Cell targetCell) {
		if (sourceCell == null) {
			// Preserve null values
			targetCell.setCellType(CellType.BLANK);
		} else {
			switch (sourceCell.getCellType()) {
			case NUMERIC:
				targetCell.setCellValue(sourceCell.getNumericCellValue());
				break;
			case STRING:
				targetCell.setCellValue(sourceCell.getStringCellValue());
				break;
			case BOOLEAN:
				targetCell.setCellValue(sourceCell.getBooleanCellValue());
				break;
			// Handle other cell types if needed
			default:
				// Do nothing or set a default value
				break;
			}
		}
	}
}