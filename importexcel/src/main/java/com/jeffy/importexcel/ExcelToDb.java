/**
 * zhangrz3
 * 2015年1月14日
 */
package com.jeffy.importexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Level;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author zhangrz3
 *
 */
public class ExcelToDb {
	
	private static Logger logger = Logger.getLogger(ExcelToDb.class);
	
	/**
	 * Excel file，need absolute path
	 */
	private String file;
	/**
	 * not need import sheet
	 */
	private List<String> exSheets;
	/**
	 * table name prefix
	 */
	private String tableNamePrefix = "";

	
	private DatabaseInfo databaseInfo;
	
	private DBTools db;

	private Connection conn;

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// --user user --pass password --url jdbcurl --file excel_file [--exclude [sheet,..]]
		if (args.length == 1 && ("--help".equals(args[1]) || "-h".equals(args[1]))) {
			showHelp();
			System.exit(0);
		}else if (!(args.length !=8 || args.length != 10)) {
			System.err.println("Wrong parameters!");
			showHelp();
			System.exit(1);
		}
		//define the variables to store parameters
		String user = null;
		String pass = null;
		String url = null;
		String file = null;
		List<String> excludeSheets = new ArrayList<String>();

		//Parse parameters
		for (int i = 0; i < args.length; i++) {
			if ("--user".equalsIgnoreCase(args[i].trim())) {
				user = args[i+1].trim();
				continue;
			}
			if ("--pass".equalsIgnoreCase(args[i].trim())) {
				pass=args[i+1].trim();
				continue;
			}
			if ("--url".equalsIgnoreCase(args[i].trim())) {
				url=args[i+1].trim();
				continue;
			}
			if ("--file".equalsIgnoreCase(args[i].trim())) {
				file = args[i+1].trim();
				continue;
			}
			if ("--exclude".equalsIgnoreCase(args[i].trim())) {
				String[] exs = args[i+1].split(",");
				for (String ex : exs) {
					excludeSheets.add(ex);
				}
				continue;
			}
		}
		
		//Check parameter
		if (user == null ) {
			System.err.println("Paramter input database user!");
			showHelp();
			System.exit(1);
		}
		if (pass == null) {
			System.err.println("Paramter input database pass!");
			showHelp();
			System.exit(1);
		}
		if (url == null) {
			System.err.println("Paramter input database jdbc URL!");
			showHelp();
			System.exit(1);
		}
		if ( file == null) {
			System.err.println("Paramter specify import file or directory!");
			showHelp();
			System.exit(1);
		}
		DatabaseInfo info = new DatabaseInfo(url, user, pass);

		File fl = new File(file);
		if (fl.isDirectory()) {
			String[] excelfiles=fl.list(new FilenameFilter() {
				@Override
				public boolean accept(File dir, String name) {
					if (name.endsWith(".xls") || name.endsWith(".xlsx")) {
						return true;
					}else {
						return false;
					}
				}
			});
			String exfile;
			for (int i = 0; i < excelfiles.length; i++) {
				if(file.endsWith("/")){
					exfile=file+excelfiles[i];
				}else {
					exfile=file +"/"+excelfiles[i];
				}
				ExcelToDb dp = new ExcelToDb(exfile,info, excludeSheets);
				try {
					dp.doImport();
				} catch (Exception e) {
					logger.error("Import："+ exfile + "failure",e);
				}finally {
					dp.release();
				}
			}
		}else if(fl.isFile()) {
			ExcelToDb dp = new ExcelToDb(file,info, excludeSheets);
			try {
				dp.doImport();
			} catch (Exception e) {
				logger.error("Import："+ file + "failure",e);
			}finally {
				dp.release();
			}
		}
		/*
		//String fileString="D:/doc/edm/Metadata/mapping/PDM Mapping-DWM-Dimension.xlsx";
		String fileString="D:/doc/edm/Metadata/Schema Columns/CRM_Column_03.xlsx";
		String tableName = "model_test";
		List<String> exlist = new ArrayList<>();
		exlist.add("Summary");
		exlist.add("Summary2");
		exlist.add("Overall");
		exlist.add("Sheet1");
		Impdp dp = new Impdp(fileString, tableName, exlist);
		dp.readtest(fileString, "CRM_Column_03");
		
		*/
	}
	public static void showHelp(){
		System.out.println("Welcome to use this tools to import Execl data to database.");
		System.out.println("Usage：");
		System.out.println("\t ExcelToDb --user user --pass password --url jdbcurl --file excel_file [--exclude [sheet,..]]");
		System.out.println("\t or");
		System.out.println("\t ExcelToDb --user user --pass password --url jdbcurl --file excel_file_dir [--exclude [sheet,..]]");
		System.out.println("Excel format must be like this, The first row is column name, the second row is column type, after that is data.");
		System.out.println("1\tid\tname\tage");
		System.out.println("2\tint\tchar(20)\tint");
		System.out.println("3\t1001\tjeffy\t25");
	}
	/**
	 * 构造方法
	 * @param file Excel文件
	 * @param exSheets 不需要导入的sheet
	 */
	public ExcelToDb(String file,DatabaseInfo info,List<String> exSheets) {
		this.file = file;
		this.databaseInfo = info;
		this.db = new DBTools(info);
		this.exSheets = exSheets;
		logger.setLevel(Level.DEBUG);
	}
	/**
	 * 获取Excel Workbook的所有Sheet名称到List中
	 * @return 包含所有Sheet名称的List对象
	 */
	public List<String> getSheets(){
		InputStream in =null;
		List<String> sheets = new ArrayList<String>();
		try{
			if (getFileExt(file).equals(ExcelMeta.EXCEL2003_POSTFIX)) {
				in = new FileInputStream(file);
				HSSFWorkbook workbook = new HSSFWorkbook(in);
				Integer numberOfSheets = workbook.getNumberOfSheets();
				for (int i = 0; i <numberOfSheets; i++) {
					sheets.add(workbook.getSheetName(i));
				}
				workbook.close();
			}
			if (getFileExt(file).equals(ExcelMeta.EXCEL2007_POSTFIX)) {
				in = new FileInputStream(file);
				XSSFWorkbook workbook = new XSSFWorkbook(in);
				Integer numberOfSheets = workbook.getNumberOfSheets();
				for (int i = 0; i <numberOfSheets; i++) {
					sheets.add(workbook.getSheetName(i));
				}
				workbook.close();
			}
		} catch (IOException e){
			logger.error("Parse "+ file+" Sheets failure!",e);
		}finally{
			try {
				in.close();
			} catch (IOException e) {
				logger.error("Close file failure!",e);
			}
		}
		
		return sheets;
	}
	/**
	 * 读取Excel数据并执行插入
	 * @throws IOException
	 */
	public void doImport() throws Exception{
		//获取所有的sheets
		List<String> allSheets = getSheets();
		//定义需要读取的sheets
		List<String> readSheets = new ArrayList<String>();
		if (exSheets != null && exSheets.size() !=0) {
			for (String sheetName : allSheets) {
				if (exSheets.contains(sheetName)) {
					continue;
				}
				readSheets.add(sheetName);
			}
		}else {
			readSheets = allSheets;
		}
		//循环读取工作表sheet
		for (String sheetName : readSheets) {
			//if (sheetName.equals("D_INV_FCT")) {
				if (getFileExt(file).equals(ExcelMeta.EXCEL2003_POSTFIX)) {
					importXLS(file, sheetName);
				}
				if (getFileExt(file).equals(ExcelMeta.EXCEL2007_POSTFIX)) {
					importXLSX(file, sheetName);
				}
			//}
		}

	}
	
	/**
	 * 读取Office 2007及以上版本的文件
	 * @param file 文件名
	 * @param sheetName 要导入的Sheet名
	 * @throws IOException 
	 */
	private void importXLSX(String file,String sheetName) throws Exception {
		//获取Excel文件的输入流
		InputStream in = new FileInputStream(file);
		//得到Excel工作簿对象
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook(in);

		//获取Excel工作表sheet
		XSSFSheet sheet = workbook.getSheet(sheetName);
		//获取工作表的第一行
		Integer firstRow = sheet.getFirstRowNum();
		//获取工作表行数
		Integer rows = sheet.getLastRowNum();
		//工作表的列数,默认先取20列，如果第20列还有值，则再往右取值。
		Integer columns = getMaxUsedColumn(sheet.getRow(firstRow),0); 
		Integer count=1;
		//循环工作表获取数据
		String tableName = tableNamePrefix+ sheetName.replaceAll("\\W", "").toUpperCase();
		for(int i = firstRow ; i < (rows+1) ; i++){
			XSSFRow row = sheet.getRow(i);
			if(row == null ){break;}
			
			if (i == firstRow) {
				//判断表是否已经存在，如果存在就不创建
				if (tableExist(tableName)) {
					i++;
					continue;
				}else {
					//获取创建表语句
					//创建表需要的信息包括，表名，列名称，列类型
					//表名使用sheet名称，列名称采用第一行的值，列类型采用第二行的值
					XSSFRow typeRow = sheet.getRow(i+1);
					i++;
					if (typeRow == null) {
						throw new Exception("Not support Excel format, the second row need specify column type!");
					}
					String sql =createTableSQL(row,typeRow,tableName,columns);
					logger.info("Create table: "+ sql);
					db.doCreate(sql);
				}
			}else {
				StringBuilder insertSQL = new StringBuilder();
				insertSQL.append("INSERT INTO "); 
				insertSQL.append(tableName);
				insertSQL.append(" VALUES (");
				List<Object> data = new ArrayList<Object>();
				boolean flag = true;
				int ept=0;
				Integer errColumn=0;
				for (int j = 0; j <columns; j++) {
					logger.debug("Import file ["+file+"] sheets<"+sheetName+"> row： " +i +" column: "+j);
					XSSFCell cell =row.getCell(j);
					String value =getValue(cell);
					//System.out.println(value);
					if (j==0) {
						if ( value== null) {
							errColumn=j;
							flag = false;
							break;
						}else {
							if (value.equals("")) {
								ept++;
							}
							data.add(value);
							insertSQL.append("?,");
						}
					}else {
						if ( value== null) {
							value="";
						}
						if (value.equals("")) {
							ept++;
						}
						data.add(value);
						insertSQL.append("?,");
					}
					
				}
				if (flag) {
					//System.out.println("总的列数：" + columns +"总的空格数："+ ept);
					if (ept == columns) {
						logger.error("File ["+file+"] sheets <"+sheetName+"> row " +i +" is empty!");
						continue;
					}else {
						insertSQL.deleteCharAt(insertSQL.length()-1);
						insertSQL.append(")");
						//System.out.println("正在插入EXCEL文件["+file+"]的工作表<"+sheetName+">第：" +i +"行");
						logger.debug("Import file ["+file+"] sheets <"+sheetName+">row " +i);
						//System.out.println(insertSQL.toString());
						//执行sql语句
						db.doUpdate(insertSQL.toString(), data);
						count++;
					}
				}else {
					logger.error("Import file ["+file+"] sheets <"+sheetName+"> row：" +i +"，column: "+errColumn+" not support format!");
				}
			}
		}
		System.out.println("Import file ["+file+"] sheets <"+sheetName+"> total： "+count+" rows.");
		workbook.close();
		in.close();
	}
	private Boolean tableExist(String name){
		String sql="SELECT table_name FROM information_schema.tables WHERE table_name=?";
		List<Object> conditions = new ArrayList<Object>();
		conditions.add(name);
		Map<String, Object> res = db.getFirst(sql, conditions);
		if (res.size()==1) {
			return true;
		}else {
			return false;
		}
	}
	
	/**
	 * 获取有数据的最大列
	 * @param row XSSFROW对象
	 * @param col 起始列
	 * @return
	 */
	private Integer getMaxUsedColumn(XSSFRow row,Integer col){
		Integer column = 0;
		if (getValue(row.getCell(col)) == null || getValue(row.getCell(col)).equals("")) {
			column =col;
		}else {
			column = getMaxUsedColumn(row, col+1);
		}
		return column;
	}
	/**
	 * 获取有数据的最大列
	 * @param row HSSFRow对象
	 * @param col 起始列
	 * @return
	 */
	private Integer getMaxUsedColumn(HSSFRow row,Integer col){
		Integer column = 0;
		if (getValue(row.getCell(col)) == null) {
			column =col;
		}else {
			column = getMaxUsedColumn(row, col+1);
		}
		return column;
	}
	/**
	 * 读取Office 2003及以下版本文件
	 * @param file Excel文件
	 * @param sheetName 要读取的Sheet名称
	 * @throws Exception 
	 */
	private void importXLS(String file,String sheetName) throws Exception {
		FileInputStream in = new FileInputStream(file);

		@SuppressWarnings("resource")
		HSSFWorkbook workbook = new HSSFWorkbook(in);

		HSSFSheet sheet=workbook.getSheet(sheetName);

		int rows = sheet.getLastRowNum();

		int firstRow = sheet.getFirstRowNum();
		int columns = getMaxUsedColumn(sheet.getRow(firstRow), 0);
		logger.debug("Import file: "+file +" sheet: " + sheetName +" column width: " + columns +" rows: "+ rows);
		int count=0;
		//循环工作表获取数据
		String tableName = tableNamePrefix+ sheetName.replaceAll("\\W", "").toUpperCase();

		for(int i = firstRow ; i < (rows+1) ; i++){
			HSSFRow row = sheet.getRow(i);
			if(row == null ){break;}

			if (i == firstRow) {
				//判断表是否已经存在，如果存在就不创建
				if (tableExist(tableName)) {
					i++;
					continue;
				}else {
					//获取创建表语句
					//创建表需要的信息包括，表名，列名称，列类型
					//表名使用sheet名称，列名称采用第一行的值，列类型采用第二行的值
					HSSFRow typeRow = sheet.getRow(i+1);
					i++;
					if (typeRow == null) {
						throw new Exception("Not support Excel format, the second row need specify column type!");
					}
					String sql =createTableSQL(row,typeRow,tableName,columns);
					logger.info("Create table: "+ sql);
					db.doCreate(sql);
				}
			}else {
				StringBuilder insertSQL = new StringBuilder();
				insertSQL.append("INSERT INTO "); 
				insertSQL.append(tableName);
				insertSQL.append(" VALUES (");
				List<Object> data = new ArrayList<Object>();
				boolean flag = true;
				int ept=0;
				Integer errColumn=0;
				for (int j = 0; j <columns; j++) {
					logger.debug("Import file ["+file+"] sheets<"+sheetName+"> row： " +i +" column: "+j);
					HSSFCell cell =row.getCell(j);
					String value =getValue(cell);
					//System.out.println(value);
					if (j==0) {
						if ( value== null) {
							errColumn=j;
							flag = false;
							break;
						}else {
							if (value.equals("")) {
								ept++;
							}
							data.add(value);
							insertSQL.append("?,");
						}
					}else {
						if ( value== null) {
							value="";
						}
						if (value.equals("")) {
							ept++;
						}
						data.add(value);
						insertSQL.append("?,");
					}
					
				}
				if (flag) {
					//System.out.println("总的列数：" + columns +"总的空格数："+ ept);
					if (ept == columns) {
						logger.error("File ["+file+"] sheets <"+sheetName+"> row " +i +" is empty!");
						continue;
					}else {
						insertSQL.deleteCharAt(insertSQL.length()-1);
						insertSQL.append(")");
						//System.out.println("正在插入EXCEL文件["+file+"]的工作表<"+sheetName+">第：" +i +"行");
						logger.debug("Import file ["+file+"] sheets <"+sheetName+">row " +i);
						//System.out.println(insertSQL.toString());
						//执行sql语句
						db.doUpdate(insertSQL.toString(), data);
						count++;
					}
				}else {
					logger.error("Import file ["+file+"] sheets <"+sheetName+"> row：" +i +"，column: "+errColumn+" not support format!");
				}
			}
		}
		System.out.println("Import file ["+file+"] sheets <"+sheetName+"> total： "+count+" rows.");
		workbook.close();
		in.close();
	}

	private String createTableSQL(HSSFRow row, HSSFRow typeRow, String tableName, int columns) {
		StringBuilder sqlBuilder = new StringBuilder();
		sqlBuilder.append("CREATE TABLE IF NOT EXISTS ");
		sqlBuilder.append(tableName);
		sqlBuilder.append(" (");
		List<String> fieldList = new ArrayList<String>();
		for(int column=0;column<columns;column++){
			String field = getValue(row.getCell(column));
			field = field.replaceAll("\\W","").toUpperCase();
			if (fieldList.contains(field)) {
				String sed =String.valueOf(System.currentTimeMillis());
				String s =sed.substring(sed.length()-4) ;
				fieldList.add(field+s);
			}else {
				fieldList.add(field);
			}
		}
		for (int i=0;i<fieldList.size();i++) {
			sqlBuilder.append("`");
			sqlBuilder.append(fieldList.get(i));
			sqlBuilder.append("` ");
			sqlBuilder.append(getValue(typeRow.getCell(i)));
			sqlBuilder.append(",");
			//sqlBuilder.append(" text,");
		}
		//sqlBuilder.append("sheet varchar(255),workbook varchar(250)");
		sqlBuilder.deleteCharAt(sqlBuilder.length()-1);
		sqlBuilder.append(") ENGINE=InnoDB DEFAULT CHARACTER SET =utf8");
		return sqlBuilder.toString();
	}
	public static String getFileExt(String path){
		if (path ==null || ExcelMeta.EMPTY.equals(path.trim())) {
			return null;
		}
		Integer len= path.lastIndexOf(ExcelMeta.POINT);
		return path.substring((len+1));
	}
	
	public static String getFileExt(File file){
		String fileString =file.getName();
		if (fileString.contains(ExcelMeta.POINT)) {
			return fileString.substring((fileString.lastIndexOf(ExcelMeta.POINT)+1), fileString.length());
		}
		return null;
	}
	private String getValue(XSSFCell xssfCell) {
		if (xssfCell == null) {
			return null;
		}
		if (xssfCell.getCellType() == XSSFCell.CELL_TYPE_BLANK) {
			return ExcelMeta.EMPTY;
		}else if (xssfCell.getCellType() == XSSFCell.CELL_TYPE_BOOLEAN) {
			return String.valueOf(xssfCell.getBooleanCellValue());
		}else if (xssfCell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
			return String.valueOf(xssfCell.getNumericCellValue());
		}else if (xssfCell.getCellType() == XSSFCell.CELL_TYPE_FORMULA) {
			return xssfCell.getStringCellValue();
		}else if (xssfCell.getCellType() == XSSFCell.CELL_TYPE_ERROR) {
			return xssfCell.getErrorCellString();
		}else {
			return xssfCell.getStringCellValue();
		}
	}
	private String getValue(HSSFCell hssfCell){
		if (hssfCell == null) {
			return null;
		}
		if (hssfCell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
			return ExcelMeta.EMPTY;
		}else if(hssfCell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
			return String.valueOf(hssfCell.getNumericCellValue());
		}else if (hssfCell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {
			return String.valueOf(hssfCell.getBooleanCellValue());
		}else if (hssfCell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
			return hssfCell.getStringCellValue();
		}else {
			return hssfCell.getStringCellValue();
		}
	}
	private String createTableSQL(XSSFRow row,XSSFRow typeRow,String tableName,Integer columns){
		StringBuilder sqlBuilder = new StringBuilder();
		sqlBuilder.append("CREATE TABLE IF NOT EXISTS ");
		sqlBuilder.append(tableName);
		sqlBuilder.append(" (");
		List<String> fieldList = new ArrayList<String>();
		for(int column=0;column<columns;column++){
			String field = getValue(row.getCell(column));
			field = field.replaceAll("\\W","").toUpperCase();
			if (fieldList.contains(field)) {
				String sed =String.valueOf(System.currentTimeMillis());
				String s =sed.substring(sed.length()-4) ;
				fieldList.add(field+s);
			}else {
				fieldList.add(field);
			}
		}
		for (int i=0;i<fieldList.size();i++) {
			sqlBuilder.append("`");
			sqlBuilder.append(fieldList.get(i));
			sqlBuilder.append("` ");
			sqlBuilder.append(getValue(typeRow.getCell(i)));
			sqlBuilder.append(",");
			//sqlBuilder.append(" text,");
		}
		//sqlBuilder.append("sheet varchar(255),workbook varchar(250)");
		sqlBuilder.deleteCharAt(sqlBuilder.length()-1);
		sqlBuilder.append(") ENGINE=InnoDB DEFAULT CHARACTER SET =utf8");
		return sqlBuilder.toString();
	} 
	public void release(){
		db.close();
	}
}
