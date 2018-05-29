package vn.iadd.excel.impl;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.function.Consumer;

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
	
	private int rowHeader;
	
	private IExcelModel modelTemplate;
	
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
	 * Read excel file to list
	 * @param file
	 * @return
	 * @throws EncryptedDocumentException
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public List<List<Object>> readFromExcel(String file) throws EncryptedDocumentException, InvalidFormatException, IOException {
		List<List<Object>> rows = new ArrayList<>();
		List<Object> cols;
		
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
			
			//int cPos = 0;
			
			cols = new ArrayList<>();
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				//cPos++;
				Cell cell = cellIterator.next();
				CellType cType = cell.getCellTypeEnum();
				if (cType == CellType.STRING) {
					cols.add(cell.getStringCellValue());
					continue;
				} else if (cType == CellType.NUMERIC) {
					if (DateUtil.isCellDateFormatted(cell)) {
						// Date
						cols.add(cell.getDateCellValue());
						//cell.setCellType(CellType.STRING);
						//cols.add(cell.getStringCellValue());
						continue;
					}
					cols.add(cell.getNumericCellValue());
					continue;
				} else if (cType == CellType.FORMULA) {
					CellValue cv = evaluator.evaluate(cell);
					if (cv.getCellTypeEnum() == CellType.STRING) {
						cols.add(cv.getStringValue());
						continue;
					} else if (cv.getCellTypeEnum() == CellType.NUMERIC) {
//						if (DateUtil.isCellDateFormatted(cell)) {
//						}
						cols.add(cv.getNumberValue());
						continue;
					} else {
						cols.add(null);
					}
				} else {
					cols.add(null);
				}
			}
			if (cols.isEmpty() || ObjectUtil.allIsNull(cols)) {
				continue;
			}
			rows.add(cols);
			
		} // End while row
		return rows;
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
		/*try {
			task.call();
		} catch (Exception e) {
			e.printStackTrace();
		}*/
	}

	/**
	 * Read all rows from excel file to list
	 * @param file String
	 * @param lst List<IExcelModel>
	 */
	private void addToList(String file, List<IExcelModel> lst) {
		if (lst == null) {
			lst = new ArrayList<>();
		}
		
		List<List<Object>> rows = null;
		try {
			rows = readFromExcel(file);
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			e.printStackTrace();
		}
		if (rows == null || rows.isEmpty()) {
			return;
		}
		int size = rows.size();
		for (int i=0; i<size; i++) {
			List<Object> cols = rows.get(i);
			IExcelModel tmp = modelTemplate.createFromFields(modelTemplate.getMapFields());
			tmp.setExcelValue(cols.toArray());
			lst.add(tmp);
		}
	}

	@Override
	public void close() throws Exception {
		if (executor != null) {
			executor.shutdown();
		}
	}
}
