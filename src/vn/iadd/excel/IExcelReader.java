package vn.iadd.excel;

import java.util.List;
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
	 * Read content of excel asynchronous. Call onDone consumer when done.
	 * @param file String
	 * @param onDone Consumer<List<IExcelModel>>
	 */
	void readAsync(String file, Consumer<List<IExcelModel>> onDone);
}
 