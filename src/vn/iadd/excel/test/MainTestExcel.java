package vn.iadd.excel.test;

import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutionException;
import java.util.function.Consumer;

import vn.iadd.excel.ExcelModel;
import vn.iadd.excel.IExcelModel;
import vn.iadd.excel.IExcelReader;
import vn.iadd.excel.impl.ExcelReaderImpl;
import vn.iadd.util.ConfigUtils;
import vn.iadd.util.Logger;

public class MainTestExcel {

	static void log(String msg) {
		Logger.log(MainTestExcel.class, msg);
	}
	
	static final String file = "output/test_20180528.xlsx";;
	
	public static void main(String[] args) throws Exception {
		int rowHeader = 2;
		String rowHeaderStr = ConfigUtils.getConfig("C_ROW_HEADER");
		if (rowHeaderStr != null) {
			try {
				rowHeader = Integer.parseInt(rowHeaderStr);
			} catch (Exception ex) {
				
			}
		}
		
		//long start, end;
		//start = System.nanoTime();
		//testReadSync(rowHeader);
		//end = System.nanoTime();
		//log("Start -> End: " + (end - start));
		//testReadAsync(rowHeader);
		testReadToMap(rowHeader);
	}
	
	static void testReadSync(int rowHeader) {
		IExcelReader reader = new ExcelReaderImpl(rowHeader, new ExcelModel());
		List<IExcelModel> lst = reader.read(file);
		try {
			reader.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		log("Done: " + lst.size());
	}
	
	static void testReadAsync(int rowHeader) throws InterruptedException, ExecutionException {
		final long start;
		start = System.nanoTime();
		IExcelReader reader = new ExcelReaderImpl(rowHeader, new ExcelModel());
		
		Consumer<List<IExcelModel>> onDone = lstObject -> {
			if (lstObject != null) {
				log("Size: " + lstObject.size());
			}
			log("I'm done here: ");
			try {
				reader.close();
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			long end = System.nanoTime();
			log("Start -> End: " + (end - start));
		};
		reader.readAsync(file, onDone);
		log("Waiting...");
	}
	
	static void testReadToMap(int rowHeader) throws InterruptedException, ExecutionException {
		final long start;
		start = System.nanoTime();
		IExcelReader reader = new ExcelReaderImpl(rowHeader, null);
		
		Consumer<List<Map<String, Object>>> onDone = lstObject -> {
			if (lstObject != null) {
				log("Size: " + lstObject.size());
			}
			log("I'm done here: ");
			try {
				reader.close();
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			long end = System.nanoTime();
			log("Start -> End: " + (end - start));
		};
		reader.readToMapAsync(file, onDone);
		log("Waiting...");
	}
}
