package example.poi;

import java.io.FileOutputStream;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
 

public class Simple {
	private static Logger LOG = Logger.getLogger(Simple.class);
	 
    public void createSimpleExcel() {
    	LOG.debug("Entra");
    	
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("Java Books");
         
        Object[][] bookData = {
                {"Head First Java", "Kathy Serria", 79},
                {"Effective Java", "Joshua Bloch", 36},
                {"Clean Code", "Robert martin", 42},
                {"Thinking in Java", "Bruce Eckel", 35},
        };
 
        int rowCount = 0;
         
        for (Object[] aBook : bookData) {
            Row row = sheet.createRow(++rowCount);
             
            int columnCount = 0;
             
            for (Object field : aBook) {
                Cell cell = row.createCell(++columnCount);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
             
        }
        
        try (FileOutputStream outputStream = new FileOutputStream("files/out/Simple.xlsx")) {
            workbook.write(outputStream);
        } catch (Exception e) {
        	LOG.error(e.getMessage(), e);
		}
    }
 
}