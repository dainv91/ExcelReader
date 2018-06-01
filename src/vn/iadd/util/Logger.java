package vn.iadd.util;

/**
 * Logger class
 * @author DaiNV
 * @since 20180529
 *
 */
public class Logger {
	
	public static void log(String msg) {
		log(Logger.class, msg);
	}
	
	public static void log(Object src, String msg) {
		if (src == null) {
			return;
		}
		log(src.getClass(), msg);
	}
	
	public static void log(Class<?> clazz, String msg) {
		System.out.println(clazz.getName() + ", -->" + Thread.currentThread().getName() + "<--->" + msg);
	}
	
	public static void main(String[] args) {
		log("test");
	}
}
