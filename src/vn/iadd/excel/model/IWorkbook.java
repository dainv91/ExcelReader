/**
 * 
 */
package vn.iadd.excel.model;

import java.util.List;
import java.util.Map;

/**
 * @author DaiNV
 *
 */
public interface IWorkbook {
	
	List<IWorksheet> getSheets();
	
	IWorksheet getSheetAt(int index);
	
	IWorksheet getSheet(String sheetName);
	
	Map<String, List<Map<String, Object>>> toMap();
}
