package vn.iadd.excel.impl;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Predicate;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import vn.iadd.excel.IExcelModel;
import vn.iadd.excel.IExcelReader;
import vn.iadd.util.Logger;
import vn.iadd.util.ObjectUtil;

public class ExcelReaderImpl implements IExcelReader {

	private static final int C_ROW_HEADER = 2;

	/**
	 * Row of header
	 */
	private int rowHeader;

	/**
	 * Object template for list field
	 */
	private IExcelModel modelTemplate;

	/**
	 * Executor for asynchronous
	 */
	ExecutorService executor = Executors.newCachedThreadPool();

	/**
	 * Constructor with object template
	 * 
	 * @param objTemplate
	 */
	public ExcelReaderImpl(IExcelModel objTemplate) {
		this(C_ROW_HEADER, objTemplate);
	}

	/**
	 * Constructor
	 * 
	 * @param rowHeader
	 * @param objTemplate
	 */
	public ExcelReaderImpl(int rowHeader, IExcelModel objTemplate) {
		this.rowHeader = rowHeader;
		this.modelTemplate = objTemplate;
	}

	void log(String msg) {
		Logger.log(this.getClass(), msg);
	}

	/**
	 * Read excel file. Call collEvent when finish read column data. Call rowEvent
	 * when finish read row data.
	 * 
	 * @param file
	 *            String
	 * @param colObj
	 * @param colEvent
	 * @param rowObj
	 * @param rowEvent
	 * @throws EncryptedDocumentException
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public <C1, R1> void readFromExcel(String file, C1 colObj, BiConsumer<C1, Object> colEvent,
			Predicate<C1> checkColsNullPre, R1 rowObj, BiConsumer<R1, C1> rowEvent)
			throws EncryptedDocumentException, InvalidFormatException, IOException {

		Workbook wb = WorkbookFactory.create(new File(file), null, true);
		Sheet sheet = wb.getSheetAt(0);
		Iterator<Row> rowIterator = sheet.rowIterator();
		FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

		int rPos = 0;
		while (rowIterator.hasNext()) {
			rPos++;
			if (rPos <= rowHeader) { // Skip header
				rowIterator.next();
				continue;
			}

			// int cPos = 0;
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				// cPos++;
				Cell cell = cellIterator.next();
				CellType cType = cell.getCellTypeEnum();
				if (cType == CellType.STRING) {
					colEvent.accept(colObj, cell.getStringCellValue());
					continue;
				} else if (cType == CellType.NUMERIC) {
					if (DateUtil.isCellDateFormatted(cell)) {
						// Date
						colEvent.accept(colObj, cell.getDateCellValue());
						continue;
					}
					colEvent.accept(colObj, cell.getNumericCellValue());
					continue;
				} else if (cType == CellType.BOOLEAN) {
					colEvent.accept(colObj, cell.getBooleanCellValue());
					continue;
				} else if (cType == CellType.FORMULA) {
					CellValue cv = evaluator.evaluate(cell);
					if (cv.getCellTypeEnum() == CellType.STRING) {
						colEvent.accept(colObj, cv.getStringValue());
						continue;
					} else if (cv.getCellTypeEnum() == CellType.BOOLEAN) {
						colEvent.accept(colObj, cv.getBooleanValue());
						continue;
					} else if (cv.getCellTypeEnum() == CellType.NUMERIC) {
						// if (DateUtil.isCellDateFormatted(cell)) {
						// }
						colEvent.accept(colObj, cv.getNumberValue());
						continue;
					} else {
						colEvent.accept(colObj, null);
					}
				} else {
					colEvent.accept(colObj, null);
				}
			}
			if (checkColsNullPre.test(colObj)) {
				continue;
			}
			rowEvent.accept(rowObj, colObj);

		} // End while row

	}

	@Override
	public List<IExcelModel> read(String file) {
		List<IExcelModel> lst = new ArrayList<>();
		addToList(file, lst);
		return lst;
	}

	@Override
	public void readAsync(String file, Consumer<List<IExcelModel>> onDone) {
		Callable<List<IExcelModel>> task = () -> {
			List<IExcelModel> lst = new ArrayList<>();
			addToList(file, lst);
			onDone.accept(lst);
			return lst;
		};
		executor.submit(task);
		/*
		 * try { task.call(); } catch (Exception e) { e.printStackTrace(); }
		 */
	}

	@Override
	public List<Map<String, Object>> readToMap(String file) {
		List<Map<String, Object>> lst = new ArrayList<>();
		addToListOfMap(file, lst);
		return lst;
	}

	@Override
	public void readToMapAsync(String file, Consumer<List<Map<String, Object>>> onDone) {
		Callable<List<Map<String, Object>>> task = () -> {
			List<Map<String, Object>> lst = new ArrayList<>();
			addToListOfMap(file, lst);
			onDone.accept(lst);
			return lst;
		};
		executor.submit(task);
	}

	/**
	 * Read all rows from excel file to list
	 * 
	 * @param file
	 *            String
	 * @param lst
	 *            List<IExcelModel>
	 */
	private void addToList(String file, List<IExcelModel> lst) {
		if (lst == null) {
			lst = new ArrayList<>();
		}

		List<List<Object>> rowObj = null;
		try {
			List<Object> colObj = new ArrayList<>();
			BiConsumer<List<Object>, Object> colEvent = (col, value) -> {
				col.add(value);
			};
			Predicate<List<Object>> checkColsNullPre = (cols) -> {
				if (cols == null || ObjectUtil.allIsNull(cols)) {
					return true;
				}
				return false;
			};

			rowObj = new ArrayList<>();
			BiConsumer<List<List<Object>>, List<Object>> rowEvent = (row, value) -> {
				List<Object> tmp = new ArrayList<>();
				tmp.addAll(value);
				row.add(tmp);
				value.clear();
			};
			readFromExcel(file, colObj, colEvent, checkColsNullPre, rowObj, rowEvent);
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			e.printStackTrace();
		}
		if (rowObj == null || rowObj.isEmpty()) {
			return;
		}
		
		int size = rowObj.size();
		for (int i = 0; i < size; i++) {
			List<Object> cols = rowObj.get(i);
			IExcelModel tmp = modelTemplate.createFromFields(modelTemplate.getMapFields());
			tmp.setExcelValue(cols.toArray());
			lst.add(tmp);
		}
	}

	/**
	 * Read all rows from excel file to list of map
	 * @param file String
	 * @param lst List<Map<String, Object>>
	 */
	private void addToListOfMap(String file, List<Map<String, Object>> lst) {
		if (lst == null) {
			lst = new ArrayList<>();
		}

		List<Map<String, Object>> rowObj = null;
		try {
			Map<String, Object> colObj = new HashMap<>();
			BiConsumer<Map<String, Object>, Object> colEvent = (col, value) -> {
				col.put(value.toString(), value.toString());
			};
			Predicate<Map<String, Object>> checkColsNullPre = (cols) -> {
				if (cols == null || ObjectUtil.allIsNull(cols.values())) {
					return true;
				}
				return false;
			};

			rowObj = new ArrayList<>();
			BiConsumer<List<Map<String, Object>>, Map<String, Object>> rowEvent = (row, value) -> {
				Map<String, Object> tmp = new HashMap<>();
				tmp.putAll(value);
				row.add(tmp);
				value.clear();
			};
			readFromExcel(file, colObj, colEvent, checkColsNullPre, rowObj, rowEvent);
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			e.printStackTrace();
		}
		if (rowObj == null || rowObj.isEmpty()) {
			return;
		}
		int size = rowObj.size();
		for (int i = 0; i < size; i++) {
			Map<String, Object> cols = rowObj.get(i);
			lst.add(cols);
		}
	}

	@Override
	public void close() throws Exception {
		if (executor != null) {
			executor.shutdown();
		}
	}
}
