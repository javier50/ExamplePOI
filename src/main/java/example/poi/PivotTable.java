package example.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PivotTable {
	private static Logger LOG = Logger.getLogger(PivotTable.class);

	public void createPivotTable() {
		LOG.debug("Entra");
		Workbook wb = new XSSFWorkbook();
		CreationHelper ch = wb.getCreationHelper();
		String safeName = WorkbookUtil.createSafeSheetName("Sample Sheet");
		XSSFSheet sheet = (XSSFSheet) wb.createSheet(safeName);

		// Lee el
		int rowNum = 0;
		List<String> colNames = null;
		try (InputStream in = new FileInputStream("files/in/baseball-salaries.csv");) {
			CSV csv = new CSV(true, ',', in);
			if (csv.hasNext()) {
				colNames = new ArrayList<String>(csv.next());
				Row row = sheet.createRow((short) 0);
				for (int i = 0; i < colNames.size(); i++) {
					String name = colNames.get(i);
					row.createCell(i).setCellValue(name);
				}
			}

			while (csv.hasNext()) {
				List<String> fields = csv.next();
				rowNum++;
				Row row = sheet.createRow((short) rowNum);
				for (int i = 0; i < fields.size(); i++) {
					/*
					 * Attempt to set as double. If that fails, set as text.
					 */
					try {
						double value = Double.parseDouble(fields.get(i));
						row.createCell(i).setCellValue(value);
					} catch (NumberFormatException ex) {
						String value = fields.get(i);
						row.createCell(i).setCellValue(value);
					}
				}
			}
		} catch (Exception e) {
			LOG.error(e.getMessage(), e);
		}
		
		// datos para la tabla dinamica
		int firstRow = sheet.getFirstRowNum();
		int lastRow = sheet.getLastRowNum();
		int firstCol = sheet.getRow(0).getFirstCellNum();
		int lastCol = sheet.getRow(0).getLastCellNum();

		
		CellReference topLeft = new CellReference(firstRow, firstCol);
		CellReference botRight = new CellReference(lastRow, lastCol - 1);

		AreaReference aref = new AreaReference(topLeft, botRight, SpreadsheetVersion.EXCEL2007);
		CellReference pos = new CellReference(firstRow + 4, lastCol + 1);
		
		XSSFPivotTable pivotTable = sheet.createPivotTable(aref, pos);
		
		pivotTable.addRowLabel(0);
		pivotTable.addRowLabel(1);
		pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 4, "Sum of " + colNames.get(4));
		
		try (FileOutputStream fileOut = new FileOutputStream("files/out/PivotTable.xlsx")) {
			wb.write(fileOut);
		} catch (Exception e) {
			LOG.error(e.getMessage(), e);
		}
	}

}