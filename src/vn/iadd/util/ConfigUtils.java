package vn.iadd.util;

import java.util.Properties;

/**
 * Config utils
 * 
 * @author DaiNV
 * @since 20180411
 */
public class ConfigUtils {

	/**
	 * Config name
	 */
	private static final String CONFIG_NAME = "/ExcelReader_20180529.properties";
	
	/**
	 * Properties
	 */
	private static Properties props;
	
	static {
		loadProp();
	}
	
	/**
	 * Load all properties
	 * @author DaiNV
	 * @since 20180411
	 */
	private static void loadProp() {
		props = new Properties();
		try {
			props.load(ConfigUtils.class.getResourceAsStream(CONFIG_NAME));
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Get config using key
	 * @param key String
	 * @return String
	 * @author DaiNV
	 * @since 20180411
	 */
	public static String getConfig(String key) {
		return props.getProperty(key);
	}
}
