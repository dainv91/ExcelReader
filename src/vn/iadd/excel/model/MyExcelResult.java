/**
 * 
 */
package vn.iadd.excel.model;

import java.util.List;

import vn.iadd.excel.IExcelResult;

/**
 * @author DaiNV
 *
 */
public final class MyExcelResult implements IExcelResult {

	//private String file;
	private IWorkbook wb;
	
	public MyExcelResult() {
		this(null);
	}
	
	public MyExcelResult(String file) {
		this(file, null, 0);
	}
	
	public MyExcelResult(String file, String password, int rowContainHeader) {
		this.wb = new MyWorkbook(file, password, rowContainHeader);
	}
	
	public static MyExcelResult newInstance() {
		return newInstance(null);
	}
	
	public static MyExcelResult newInstance(String file) {
		return new MyExcelResult(file);
	}
	
	/* (non-Javadoc)
	 * @see vn.iadd.excel.IExcelResult#getWorkbook()
	 */
	@Override
	public IWorkbook getWorkbook() {
		// TODO Auto-generated method stub
		return this.wb;
	}

	public static void main(String[] args) {
		final String file = "C:\\Users\\DaiNV\\Documents\\read_20181203.xlsx";
		IExcelResult result = MyExcelResult.newInstance(file);
		List<?> obj = result.getWorkbook().getSheetAt(1).getRows();
		for (Object o: obj) {
			System.out.println(o);
		}
	}
}
