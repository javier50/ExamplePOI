package example.poi;

import org.apache.log4j.Logger;

public class Main {
	private static Logger LOG = Logger.getLogger(Main.class);
	
    public static void main(String[] args) throws Exception {
    	LOG.debug("Entra");
    	Simple simpleExcel = new Simple();
    	simpleExcel.createSimpleExcel();
    	PivotTable pivotTable = new PivotTable();
    	pivotTable.createPivotTable();
	}
}
