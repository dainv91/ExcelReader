package vn.iadd.excel.impl;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.function.Consumer;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import vn.iadd.excel.IExcelReader;

public abstract class ExcelReaderV1Impl implements IExcelReader {
	
	void log(String msg) {
		log(this.getClass(), msg);
	}
	
	void log(Class<?> clazz, String msg) {
		System.out.println(clazz.getName() + ", -->" + Thread.currentThread().getName() + "<--->" + msg);
	}
	
	//@Override
	public <T> List<T> read1(String file) {
		return null;
	}

	//@Override
	public <T> void readAsync1(String file, Consumer<List<T>> onDone) {
		ExecutorService executor = Executors.newCachedThreadPool();
		Callable<List<T>> task = () -> {
			List<T> lst = new ArrayList<>();
			log(this.getClass(), "I'm executing...");
			test(file, lst);
			onDone.accept(lst);
			return lst;
		};
		executor.submit(task);
//		try {
//			onDone.accept(task.call());
//		} catch (Exception e) {
//			e.printStackTrace();
//		}
	}
	
	/**
     * See org.xml.sax.helpers.DefaultHandler javadocs
     */
    private static class SheetHandler extends DefaultHandler {
        private final SharedStringsTable sst;
        private String lastContents;
        private boolean nextIsString;
        private boolean inlineStr;
        private final LruCache<Integer,String> lruCache = new LruCache<>(50);

        private static class LruCache<A,B> extends LinkedHashMap<A, B> {
            /**
			 * 
			 */
			private static final long serialVersionUID = 1L;
			private final int maxEntries;

            public LruCache(final int maxEntries) {
                super(maxEntries + 1, 1.0f, true);
                this.maxEntries = maxEntries;
            }

            @Override
            protected boolean removeEldestEntry(final Map.Entry<A, B> eldest) {
                return super.size() > maxEntries;
            }
        }

        private SheetHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

        @Override
        public void startElement(String uri, String localName, String name,
                                 Attributes attributes) throws SAXException {
            // c => cell
            if(name.equals("c")) {
                // Print the cell reference
                System.out.print(attributes.getValue("r") + " - ");
                // Figure out if the value is an index in the SST
                String cellType = attributes.getValue("t");
                nextIsString = cellType != null && cellType.equals("s");
                inlineStr = cellType != null && cellType.equals("inlineStr");
            }
            // Clear contents cache
            lastContents = "";
        }

        @Override
        public void endElement(String uri, String localName, String name)
                throws SAXException {
            // Process the last contents as required.
            // Do now, as characters() may be called more than once
            if(nextIsString) {
                Integer idx = Integer.valueOf(lastContents);
                lastContents = lruCache.get(idx);
                if (lastContents == null && !lruCache.containsKey(idx)) {
                    lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                    lruCache.put(idx, lastContents);
                }
                nextIsString = false;
            }

            // v => contents of a cell
            // Output after we've seen the string contents
            if(name.equals("v") || (inlineStr && name.equals("c"))) {
                System.out.println(lastContents);
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) throws SAXException { // NOSONAR
            lastContents += new String(ch, start, length);
        }
    }
	
	public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException, ParserConfigurationException {
        XMLReader parser = SAXHelper.newXMLReader();
        ContentHandler handler = new SheetHandler(sst);
        parser.setContentHandler(handler);
        return parser;
    }
	
	private void test(String file, List<?> lst) throws InvalidOperationException, IOException, OpenXML4JException, SAXException, ParserConfigurationException {
		try (OPCPackage pkg = OPCPackage.open(file, PackageAccess.READ)) {
			XSSFReader r = new XSSFReader(pkg);
			SharedStringsTable sst = r.getSharedStringsTable();
			XMLReader parser = fetchSheetParser(sst);
			// process the first sheet
            try (InputStream sheet = r.getSheetsData().next()) {
				InputSource sheetSource = new InputSource(sheet);
				parser.parse(sheetSource);
			}
		}
	}

	@Override
	public void close() throws Exception {
		// TODO Auto-generated method stub
		
	}
}
