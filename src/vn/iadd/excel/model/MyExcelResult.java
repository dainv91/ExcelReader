/**
 * 
 */
package vn.iadd.excel.model;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.util.List;
import java.util.Map;

import vn.iadd.excel.IExcelResult;
import vn.iadd.util.Logger;

/**
 * @author DaiNV
 *
 */
public final class MyExcelResult implements IExcelResult {

	// private String file;
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
	
	public MyExcelResult(InputStream fileInputStream, String password, int rowContainHeader) {
		this.wb = new MyWorkbook(fileInputStream, password, rowContainHeader);
	}

	public static MyExcelResult newInstance() {
		return newInstance("");
	}

	public static MyExcelResult newInstance(String file) {
		return new MyExcelResult(file);
	}
	
	public static MyExcelResult newInstance(InputStream fileInputStream) {
		return new MyExcelResult(fileInputStream, null, 0);
	}

	/*
	 * (non-Javadoc)
	 * 
	 * @see vn.iadd.excel.IExcelResult#getWorkbook()
	 */
	@Override
	public IWorkbook getWorkbook() {
		// TODO Auto-generated method stub
		return this.wb;
	}

	public static void serializeTo(Object obj, String fileName) {
		try {
			FileOutputStream fos = new FileOutputStream(fileName);
			ObjectOutputStream oos = new ObjectOutputStream(fos);
			oos.writeObject(obj);
			oos.close();
			fos.close();
			Logger.log("Serialized HashMap data is saved in: " + fileName);
		} catch (IOException ioe) {
			ioe.printStackTrace();
		}
	}

	@SuppressWarnings("unchecked")
	public static <T> T serializeFrom(Class<T> clazz, String fileName) {
		T obj = null;
		try {
			FileInputStream fis = new FileInputStream(fileName);
			ObjectInputStream ois = new ObjectInputStream(fis);
			obj = (T) ois.readObject();
			ois.close();
			fis.close();
		} catch (IOException ioe) {
			ioe.printStackTrace();
		} catch (ClassNotFoundException c) {
			Logger.log("Class not found");
			c.printStackTrace();
		}
		return obj;
	}

	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static void main(String[] args) {
		final String file = "C:\\Users\\DaiNV\\Documents\\read_20181203.xlsx";
		IExcelResult result = MyExcelResult.newInstance(file);
		Object obj = result.getWorkbook().toMap();
		
		final String outFile = "C:\\Users\\DaiNV\\Desktop\\output.ser";
		serializeTo(obj, outFile);
		Class<?> clazz = Map.class;
		Map<String, List<Map<String, Object>>> deserialize = (Map) serializeFrom(clazz, outFile);
		System.out.println(deserialize);
		System.out.println(deserialize.get("data_sheet_1").get(1).get("lastName"));
	}
}
