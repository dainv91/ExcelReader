package vn.iadd.util;

import java.beans.BeanInfo;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

/**
 * Object utils
 * 
 * @author DaiNV
 * @since 20180511
 */
public class ObjectUtil {

	/**
	 * Check all element is null
	 * 
	 * @param lst
	 *            List<Object>
	 * @return boolean
	 * @author DaiNV
	 * @since 20180529
	 */
	public static <T>boolean allIsNull(Collection<T> lst) {
		if (lst == null) {
			return true;
		}
		for (T obj : lst) {
			if (obj != null) {
				return false;
			}
		}
		return true;
	}

	/**
	 * Get all fields of object
	 * 
	 * @param obj
	 *            Object
	 * @return List<Field
	 * @author DaiNV
	 * @since 20180528
	 */
	public static List<Field> getFields(Object obj) {
		List<Field> fields = new ArrayList<>();
		Class<?> clazz = obj.getClass();
		while (clazz != Object.class) {
			fields.addAll(Arrays.asList(clazz.getDeclaredFields()));
			clazz = clazz.getSuperclass();
		}
		return fields;
	}

	/**
	 * Copy all properties of object.
	 * 
	 * @param src
	 *            Object
	 * @param dest
	 *            Object
	 * @author KBSDaiNV
	 * @since 20170726
	 */
	public static void copyProperties(Object src, Object dest) {
		if (src == null || dest == null) {
			return;
		}
		Class<? extends Object> fromClass = src.getClass();
		Class<? extends Object> toClass = dest.getClass();
		final String cName = "class";
		try {
			BeanInfo fromBean = Introspector.getBeanInfo(fromClass);
			BeanInfo toBean = Introspector.getBeanInfo(toClass);

			PropertyDescriptor[] toPd = toBean.getPropertyDescriptors();
			List<PropertyDescriptor> fromPd = Arrays.asList(fromBean.getPropertyDescriptors());
			for (PropertyDescriptor propertyDescriptor : toPd) {
				propertyDescriptor.getDisplayName();
				PropertyDescriptor pd = fromPd.get(fromPd.indexOf(propertyDescriptor));
				boolean sameMethod = pd.getDisplayName().equals(propertyDescriptor.getDisplayName());
				if (sameMethod && !pd.getDisplayName().equals(cName)) {
					if (propertyDescriptor.getWriteMethod() != null) {
						if (pd.getReadMethod() == null) {
							continue;
						}
						try {
							Object args = pd.getReadMethod().invoke(src);
							propertyDescriptor.getWriteMethod().invoke(dest, args);
						} catch (Exception ex) {
							// Skip properties
						}
					}
				}

			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// public static void getAllProperties(BaseExcelModel obj) throws Exception {
	// if (obj == null) {
	// return;
	// }
	// Map<String, Field> mapFields = new HashMap<>();
	// for (Field f : getFields(obj)) {
	// boolean isFinal = false;
	// isFinal = isFinal || java.lang.reflect.Modifier.isFinal(f.getModifiers());
	// mapFields.put(f.getName(), f);
	// if (isFinal) {
	// continue;
	// }
	// // f.set("");
	// f.setAccessible(true);
	// }
	// /*obj.set("name", "Dai1");
	// obj.set("num", 2);
	// obj.set("mapColumnNameToIndex", new HashMap<String, Integer>() {
	// {
	// put("key1", 5);
	// put("key2", 10);
	// }
	// });*/
	// //System.out.println(obj);
	// }

	public static void main(String[] args) throws Exception {
		// long start, end;
		// start = System.nanoTime();
		// ExcelModel objTemplate = new ExcelModel();
		// for (int i=0; i<1000; i++) {
		// ExcelModel tmp = new ExcelModel(objTemplate.getMapFields());
		// tmp.setExcelValue(new Object[] {"VDN_" + i, "dainv_@gmail.com_" + i, "dainv_"
		// + i, "Hpt123456"});
		// getAllProperties(tmp);
		// }
		// end = System.nanoTime();
		// System.out.println(end - start);
	}
}
