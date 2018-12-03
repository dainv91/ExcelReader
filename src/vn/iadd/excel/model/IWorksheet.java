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
public interface IWorksheet {

	String getSheetName();
	
	int getRowContainHeader();
	
	Map<String, Integer> getHeaderMap();
	
	Map<String, Object> getRow(int row);
	
	List<Map<String, Object>> getRows();
}
