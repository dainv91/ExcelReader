package vn.iadd.excel;

import java.lang.reflect.Field;
import java.util.Date;
import java.util.Map;

public class ExcelModel extends BaseExcelModel {

	/**
	 * Serial version
	 */
	private static final long serialVersionUID = 1L;
	
	private String fullName, email, login, pass;
	private Double actKey;
	private Date createdDate;

	public ExcelModel() {
		
	}
	
	public ExcelModel(Map<String, Field> fields) {
		super(fields);
	}
	
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

	@Override
	public IExcelModel createFromFields(Map<String, Field> mapFields) {
		ExcelModel obj = new ExcelModel(mapFields);
		return obj;
	}

	public String getFullName() {
		return fullName;
	}

	public void setFullName(String fullName) {
		this.fullName = fullName;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	public String getLogin() {
		return login;
	}

	public void setLogin(String login) {
		this.login = login;
	}

	public String getPass() {
		return pass;
	}

	public void setPass(String pass) {
		this.pass = pass;
	}

	public Double getActKey() {
		return actKey;
	}

	public void setActKey(Double actKey) {
		this.actKey = actKey;
	}
	
	public Date getCreatedDate() {
		return createdDate;
	}

	public void setCreatedDate(Date createdDate) {
		this.createdDate = createdDate;
	}

	@Override
	public String toString() {
		return "ExcelModel [fullName=" + fullName + ", email=" + email + ", login=" + login + ", pass=" + pass
				+ ", actKey=" + actKey + ", createdDate=" + createdDate + "]";
	}
}
