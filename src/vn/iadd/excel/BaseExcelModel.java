package vn.iadd.excel;

import java.io.Serializable;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import vn.iadd.util.ObjectUtil;

/**
 * BaseExcelModel
 * 
 * @author DaiNV
 * @since 20180529
 */
public abstract class BaseExcelModel implements IExcelModel, Serializable {

	/**
	 * Serial version
	 */
	private static final long serialVersionUID = 1L;

	/**
	 * Map column name - index
	 */
	private Map<String, Integer> mapColumnNameToIndex = new HashMap<>();
	
	/**
	 * Map index - column name
	 */
	private Map<Integer, String> mapIndexToColumnName = new HashMap<>();

	/**
	 * Map property name - field
	 */
	private final Map<String, Field> mapFields = new HashMap<>();

	/**
	 * Default constructor
	 */
	public BaseExcelModel() {
		initFields();
		mapColumnWithIndex();
	}

	/**
	 * Constructor with parameter
	 * @param fields Map<String, Field>
	 */
	public BaseExcelModel(Map<String, Field> fields) {
		if (fields != null && !fields.isEmpty()) {
			//mapFields.clear();
			this.mapFields.putAll(fields);
		}
		mapColumnWithIndex();
	}

	public boolean containColumn(String colName) {
		return mapColumnNameToIndex.containsKey(colName);
	}

	/**
	 * Add column name with index in excel file
	 * @param colName String
	 * @param index int
	 */
	public void addColumnIndex(String colName, int index) {
		Integer i = Integer.valueOf(index);
		mapColumnNameToIndex.put(colName, i);
		mapIndexToColumnName.put(i, colName);
	}

	public int getColumnIndex(String colName) {
		return mapColumnNameToIndex.get(colName).intValue();
	}

	/**
	 * Get map fields using reflection.
	 */
	private void initFields() {
		initFields(ObjectUtil.getFields(this));
	}
	
	/**
	 * Get map fields using another object template.
	 * @param fields
	 */
	private void initFields(List<Field> fields) {
		if (fields == null || fields.isEmpty()) {
			return;
		}
		mapFields.clear();
		for (Field f : fields) {
			boolean isFinal = false;
			isFinal = isFinal || java.lang.reflect.Modifier.isFinal(f.getModifiers());
			mapFields.put(f.getName(), f);
			if (isFinal) {
				continue;
			}
			f.setAccessible(true);
		}
	}
	
	@Override
	public Map<String, Field> getMapFields() {
		return mapFields;
	}
	
	public Map<String, Integer> getMapColumnNameToIndex() {
		return mapColumnNameToIndex;
	}
	
	public void setMapColumnNameToIndex(Map<String, Integer> mapColumnNameToIndex) {
		this.mapColumnNameToIndex = mapColumnNameToIndex;
	}

	public void set(String propertyName, Object value) {
		if (!mapFields.containsKey(propertyName)) {
			return;
		}
		try {
			mapFields.get(propertyName).set(this, value);
		} catch (IllegalArgumentException | IllegalAccessException e) {
			e.printStackTrace();
		}
	}

	public Object get(String propertyName) {
		if (!mapFields.containsKey(propertyName)) {
			return null;
		}
		return mapFields.get(propertyName);
	}
	
	@Override
	public void setExcelValue(Object[] arr) {
		if (arr == null || arr.length == 0) {
			return;
		}
		for (int i = 0; i < arr.length; i++) {
			Integer col = Integer.valueOf(i + 1);
			Object value = arr[i];
			
			if (!mapIndexToColumnName.containsKey(col)) {
				continue;
			}
			String colName = mapIndexToColumnName.get(col);
			this.set(colName, value);
		}
	}
}
