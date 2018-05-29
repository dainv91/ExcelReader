package vn.iadd.excel;

import java.lang.reflect.Field;
import java.util.Map;

public interface IExcelModel {
	
	IExcelModel createFromFields(Map<String, Field> mapFields);
	
	Map<String, Field> getMapFields();
	
	void mapColumnWithIndex();
	
	void setExcelValue(Object[] arr);
}
