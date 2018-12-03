/**
 * 
 */
package vn.iadd.excel.model;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * @author DaiNV
 *
 */
public final class MyWorksheet implements IWorksheet {

	private int rowContainHeader;
	private Sheet sheet;
	
	public MyWorksheet(Sheet sheet) {
		this(sheet, 0);
	}
	
	public MyWorksheet(Sheet sheet, int rowContainHeader) {
		this.sheet = sheet;
		this.rowContainHeader = rowContainHeader;
	}
	
	/**
	 * Get value of cell
	 * @param cell Cell
	 * @param evaluator FormulaEvaluator
	 * @return Object
	 */
	private Object getCellValue(Cell cell, FormulaEvaluator evaluator) {
		Object value = null;
		if (cell == null) {
			return value;
		}
		CellType cType = cell.getCellTypeEnum();
		if (cType == CellType.STRING) {
			value = cell.getStringCellValue();
		} else if (cType == CellType.NUMERIC) {
			if (DateUtil.isCellDateFormatted(cell)) {
				// Date
				value = cell.getDateCellValue();
			} else {
				value = cell.getNumericCellValue();	
			}
		} else if (cType == CellType.BOOLEAN) {
			value = cell.getBooleanCellValue();
		} else if (cType == CellType.FORMULA) {
			CellValue cv = evaluator.evaluate(cell);
			if (cv.getCellTypeEnum() == CellType.STRING) {
				value = cv.getStringValue();
			} else if (cv.getCellTypeEnum() == CellType.BOOLEAN) {
				value = cv.getBooleanValue();
			} else if (cv.getCellTypeEnum() == CellType.NUMERIC) {
				// if (DateUtil.isCellDateFormatted(cell)) {
				// }
				value = cv.getNumberValue();
			} else {
				value = null;
			}
		}
		return value;
	}
	
	/* (non-Javadoc)
	 * @see vn.iadd.excel.model.IWorksheet#getSheetName()
	 */
	@Override
	public String getSheetName() {
		return sheet.getSheetName();
	}

	/* (non-Javadoc)
	 * @see vn.iadd.excel.model.IWorksheet#getRowContainHeader()
	 */
	@Override
	public int getRowContainHeader() {
		return rowContainHeader;
	}

	/* (non-Javadoc)
	 * @see vn.iadd.excel.model.IWorksheet#getHeaderMap()
	 */
	@Override
	public Map<String, Integer> getHeaderMap() {
		FormulaEvaluator evaluator = this.sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		Map<String, Integer> map = new HashMap<>();
		Row row = sheet.getRow(rowContainHeader);
		short last = row.getLastCellNum();
		for (int i=0; i<last; i++) {
			final Cell c = row.getCell(i);
			final Object value = getCellValue(c, evaluator);
			if (value == null) {
				continue;
			}
			final String strValue = value.toString();
			map.put(strValue, Integer.valueOf(i));
		}
		//System.out.println(map);
		return map;
	}

	/* (non-Javadoc)
	 * @see vn.iadd.excel.model.IWorksheet#getRow(int)
	 */
	@Override
	public Map<String, Object> getRow(int rNum) {
		FormulaEvaluator evaluator = this.sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		Map<String, Object> map = new HashMap<>();
		Row row = sheet.getRow(rNum);
		if (row == null) {
			return map;
		}
		
		Map<String, Integer> mapHeader = getHeaderMap();
		for (String header: mapHeader.keySet()) {
			Integer col = mapHeader.get(header);
			if (col == null) {
				continue;
			}
			final Cell c = row.getCell(col);
			final Object value = getCellValue(c, evaluator);
			map.put(header, value);
		}
		return map;
	}

	/* (non-Javadoc)
	 * @see vn.iadd.excel.model.IWorksheet#getRows()
	 */
	@Override
	public List<Map<String, Object>> getRows() {
		List<Map<String, Object>> lst = new ArrayList<>();
		int lastRowNum = sheet.getLastRowNum();
		for (int i=0; i<=lastRowNum; i++) {
			if (i <= rowContainHeader) {
				continue;
			}
			Map<String, Object> row = getRow(i);
			lst.add(row);
		}
		return lst;
	}

}
