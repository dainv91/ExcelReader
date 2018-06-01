# ExcelReader
This project using Apache POI read excel file to list object

## Getting started

Git clone the project on your local machine and add it to your workspace (I'm using Eclipse)

### Prerequisites

For running this, you will need
 - java 1.8
 - Apache POI (https://poi.apache.org/download.html)

## How to

1. Create class extends **BaseExcelModel** (for example class **ExcelModel**). It should implements method **mapColumnWithIndex**.
If you only need read each row to Map<String, Object> (readToMap, readToMapAsync), skip step 1, and pass **null** at 2nd parameter at step 2.
```java
@Override
public void mapColumnWithIndex() {
	int pos = 1;
	addColumnIndex("No.", pos++);
	addColumnIndex("fullName", pos++);
	addColumnIndex("email", pos++);
	addColumnIndex("login", pos++);
	addColumnIndex("pass", pos++);
	addColumnIndex("actKey", pos++);
	addColumnIndex("createdDate", pos++);
}
```
- And method **createFromFields**
```java
@Override
public IExcelModel createFromFields(Map<String, Field> mapFields) {
	ExcelModel obj = new ExcelModel(mapFields);
	return obj;
}
```

2. Create new instance of **IExcelReader**. It takes **rowHeader** and an instance of **IExcelModel** as parameters. For example **ExcelReaderImpl**
```java
IExcelReader reader = new ExcelReaderImpl(rowHeader, new ExcelModel());
```
3. Now you can use these function of **IExcelReader**
```java
List<IExcelModel> read(String file);	
void readAsync(String file, Consumer<List<IExcelModel>> onDone);
List<Map<String, Object>> readToMap(String file);
void readToMapAsync(String file, Consumer<List<Map<String, Object>>> onDone);
```

*You can check sample in class* **MainTestExcel**
```java
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
```