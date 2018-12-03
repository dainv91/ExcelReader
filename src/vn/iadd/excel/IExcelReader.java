package vn.iadd.excel;

import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

/**
 * IExcelReader
 * @author DaiNV
 * @since 20180529
 */
public interface IExcelReader extends AutoCloseable {
	
	/**
	 * Read content of excel file to list
	 * @param file String
	 * @return List<IExcelModel>
	 */
	List<IExcelModel> read(String file);
	
	/**
	 * Read content of all sheet
	 * @param file String
	 * @return IExcelResult
	 */
	IExcelResult readAllSheet(String file);
	
	/**
	 * Read content of excel asynchronous. Call onDone consumer when finish.
	 * @param file String
	 * @param onDone Consumer<List<IExcelModel>>
	 */
	void readAsync(String file, Consumer<List<IExcelModel>> onDone);
	
	/**
	 * Read content of excel file to list of map object
	 * @param file String
	 * @return List<Map<String, Object>>
	 */
	List<Map<String, Object>> readToMap(String file);
	
	/**
	 * Read content of excel file to list of map object asynchronous. Call onDone consumer when finish.
	 * @param file String
	 * @param onDone Consumer<List<Map<String, Object>>>
	 */
	void readToMapAsync(String file, Consumer<List<Map<String, Object>>> onDone);
}
 