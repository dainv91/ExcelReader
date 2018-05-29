package vn.iadd.excel;

import java.util.List;
import java.util.function.Consumer;

public interface IExcelReader extends AutoCloseable {
	
	List<IExcelModel> read(String file);
	
	void readAsync(String file, Consumer<List<IExcelModel>> onDone);
}
 