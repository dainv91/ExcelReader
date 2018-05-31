package vn.iadd.excel;

import java.lang.reflect.Field;
import java.util.Map;

/**
 * IExcelModel
 * @author DaiNV
 * @since 20180529
 */
public interface IExcelModel {
	
	/**
	 * Copy fields from another object instead of reflection.
	 * @param mapFields Map<String, Field>
	 * @return IExcelModel
	 */
	IExcelModel createFromFields(Map<String, Field> mapFields);
	
	/**
	 * Get map name - field
	 * @return Map<String, Field>
	 */
	Map<String, Field> getMapFields();
	
	/**
	 * Map column name with index
	 */
	void mapColumnWithIndex();
	
	/**
	 * Set value from excel
	 * @param arr Object[]
	 */
	void setExcelValue(Object[] arr);
}
