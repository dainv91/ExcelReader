/**
 * 
 */
package vn.iadd.excel.model;

import java.util.List;

/**
 * @author DaiNV
 *
 */
public interface IWorkbook {
	
	List<IWorksheet> getSheets();
	
	IWorksheet getSheetAt(int index);
	
	IWorksheet getSheet(String sheetName);
}
