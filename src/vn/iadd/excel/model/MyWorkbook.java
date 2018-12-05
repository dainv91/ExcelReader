/**
 * 
 */
package vn.iadd.excel.model;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * @author DaiNV
 *
 */
public final class MyWorkbook implements IWorkbook {

	private int rowContainHeader;
	//private String file, password;
	
	private static final Map<String, Integer> mapSheetName = new HashMap<>();
	private static final Map<String, Sheet> mapSheets = new HashMap<>();
	
	private Workbook wb;
	
	public MyWorkbook() {
		// TODO Auto-generated constructor stub
		this(null);
	}
	
	public MyWorkbook(String file) {
		this(file, null);
	}
	
	public MyWorkbook(String file, String password) {
		this(file, password, 0);
	}
	
	public MyWorkbook(String file, String password, int rowContainHeader) {
		this.rowContainHeader = rowContainHeader;
		//this.file = file;
		//this.password = password;
		try {
			read0(file, password);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			clearOldData();
			mapSheetName.clear();
			mapSheets.clear();
		}
	}
	
	public MyWorkbook(InputStream fileInputStream, String password, int rowContainHeader) {
		this.rowContainHeader = rowContainHeader;
		//this.password = password;
		try {
			read0(fileInputStream, password);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			clearOldData();
		}
	}
	
	private void clearOldData() {
		this.wb = null;
	}
	
	private void read0(final String fileName, final String password) throws EncryptedDocumentException, InvalidFormatException, IOException {
		if (fileName == null || fileName.trim().isEmpty()) {
			return;
		}
		this.wb = WorkbookFactory.create(new File(fileName), password, true);
		readWorkbook0();
	}
	
	private void read0(InputStream in, final String password) throws EncryptedDocumentException, InvalidFormatException, IOException {
		this.wb = WorkbookFactory.create(in, password);
		//this.wb = WorkbookFactory.create(new File(file), password, true);
		readWorkbook0();
	}
	
	private void readWorkbook0() {
		if (wb == null) {
			return;
		}
		int sheetCount = wb.getNumberOfSheets();
		for (int i=0; i<sheetCount; i++) {
			final Sheet sheet = wb.getSheetAt(i);
			final String sheetName = sheet.getSheetName();
			//Logger.log("Sheet name: " + sheetName + "==" + i);
			mapSheetName.put(sheetName, i);
			mapSheets.put(sheetName, sheet);
		}
	}
	
	/* (non-Javadoc)
	 * @see vn.iadd.excel.model.IWorkbook#getSheets()
	 */
	@Override
	public List<IWorksheet> getSheets() {
		List<IWorksheet> lst = new ArrayList<>();
		for (Sheet s: mapSheets.values()) {
			IWorksheet sheet = new MyWorksheet(s);
			lst.add(sheet);
		}
		return lst;
	}

	/* (non-Javadoc)
	 * @see vn.iadd.excel.model.IWorkbook#getSheetAt(int)
	 */
	@Override
	public IWorksheet getSheetAt(int index) {
		Sheet s = wb.getSheetAt(index);
		IWorksheet sheet = new MyWorksheet(s, rowContainHeader);
		return sheet;
	}

	/* (non-Javadoc)
	 * @see vn.iadd.excel.model.IWorkbook#getSheet(java.lang.String)
	 */
	@Override
	public IWorksheet getSheet(String sheetName) {
		Sheet s = mapSheets.get(sheetName);
		IWorksheet sheet = new MyWorksheet(s);
		return sheet;
	}

	/* (non-Javadoc)
	 * @see vn.iadd.excel.model.IWorkbook#toMap()
	 */
	@Override
	public Map<String, List<Map<String, Object>>> toMap() {
		Map<String, List<Map<String, Object>>> mapData = new HashMap<>();
		List<IWorksheet> sheets = getSheets();
		for (IWorksheet sheet: sheets) {
			final String sName = sheet.getSheetName();
			final List<Map<String, Object>> rows = sheet.getRows();
			mapData.put(sName, rows);
		}
		return mapData;
	}
}
