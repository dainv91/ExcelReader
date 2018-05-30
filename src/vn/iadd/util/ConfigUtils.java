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
		loadProp(CONFIG_NAME);
	}

	/**
	 * Load all properties
	 * 
	 * @param String configFile
	 * @author DaiNV
	 * @since 20180411
	 */
	public static void loadProp(String configFile) {
		props = new Properties();
		try {
			props.load(ConfigUtils.class.getResourceAsStream(configFile));
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Get config using key
	 * 
	 * @param key
	 *            String
	 * @return String
	 * @author DaiNV
	 * @since 20180411
	 */
	public static String getConfig(String key) {
		return props.getProperty(key);
	}
}
