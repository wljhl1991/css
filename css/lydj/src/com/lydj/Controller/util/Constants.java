package com.lydj.Controller.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.lydj.Controller.UploadController;

/**
 * 系统常量
 * 
 * @author ldj
 * 
 */
public class Constants {
	 private static Logger logger = LoggerFactory.getLogger(UploadController.class);
	/**
	 * 模版位置
	 */
	public static String TEMPLET = "templet"+File.separator+"商务楼宇工作站情况汇总表.xlsx";
	/**
	 * 打印模版位置
	 */
	public static String PRINTTEMPLET = "templet"+File.separator+"打印模版.xlsx";
	/**
	 * 采集表模版位置（区县+街道）
	 */
	public static String TEMPLETDATE = "templet"+File.separator+"工作站信息采集表.xlsx";
	/**
	 * 数据（区县+街道）
	 */
	public static String DATATOWNSTREET = "templet"+File.separator+"北京市行政区域代码(2015).xlsx";
	/**
	 * 生成文档位置（区县+街道）
	 */
	public static String GENEDOWNLOAD = "upload"+File.separator;
	

	/**
	 * 生成文档位置（汇总）
	 */
	public static String WRITEPATH = "download"+File.separator+"";
	
	/**
	 * 页面随机字符串长度
	 */
	public static int RONDOMLENGTH = 20;
	
	/**
	 * 单Sheet页复制所有样式 开关 (严重影响时间)
	 */
	public static boolean SWITCHSTYLECOPY = false ;
	
	/**
	 * 单Sheet页复制超时 开关
	 */
	public static boolean SWITCHOVERTIME = true ;
	/**
	 * 单Sheet页复制超时次数
	 */
	public static int OVERTIMECOUNT = 5 ;
	
	/**
	 * 单Sheet页复制超时时间(毫秒)
	 * 时间间隔
	 */
	public static long OVERTIME = 1000 * 4 ;
	
	
	/**
	 * 写入超时 开关
	 */
	public static boolean SWITCHWRITE = true ;
	/**
	 * 最大写入时间
	 */
	public static long OVERWRITETIME = 1000 * 100 ;
	
	
	// 汉字 数字
	public static Map<String, String> totalMap = new HashMap<String, String>();
	// 数字 汉字
//	public static Map<String, String> basicMap = new HashMap<String, String>();

	// 数字 汉字
	public static Map<String, String> totalMapRec = new HashMap<String, String>();
	// 汉字 数字
//	public static Map<String, String> basicMapRec = new HashMap<String, String>();

	// 打印模版读取
	// 汉字 数字
	public static Map<String, String> p1Map = new HashMap<String, String>();
	// 数字 汉字
	public static Map<String, String> p1MapRec = new HashMap<String, String>();
	
	//原来打印版为4页  打算用map存储打印模版 现在已废弃
	// 汉字 数字
	public static Map<String, String> p2Map = new HashMap<String, String>();
	// 数字 汉字
	public static Map<String, String> p2MapRec = new HashMap<String, String>();
	// 汉字 数字
	public static Map<String, String> p3Map = new HashMap<String, String>();
	// 数字 汉字
	public static Map<String, String> p3MapRec = new HashMap<String, String>();
	// 汉字 数字
	public static Map<String, String> p4Map = new HashMap<String, String>();
	// 数字 汉字
	public static Map<String, String> p4MapRec = new HashMap<String, String>();

	//最后 累计 公式等 
	// 数字 汉字
	public static Map<Integer, String> forMap = new HashMap<Integer, String>();
	
	// 汉字 数字
	public static Map<String, Integer> forMapRec = new HashMap<String, Integer>();

	//区县 和街道数据
	public static Map<String,String> allTown = new HashMap<String, String>();
	public static Map<String,Map<String,String>> allStreet = new HashMap<String, Map<String,String>>();
	static {
		getAllMap();
	}

	/**
	 * 服务器绝对路径
	 */
	public static String getRealServerPath() {
		String path = UploadController.class.getResource(""+File.separator+"")+"";
		logger.debug("当前path："+path);
		int lastIndex = path.lastIndexOf("WEB-INF");
		String rootPath = path.substring(6, lastIndex).replace("\\", File.separator);
		logger.debug("当前服务器绝对路径："+rootPath);
		return rootPath;
	}

	/**
	 * 模版位置绝对路径
	 */
	public static String getTempletPath() {
		String path = getRealServerPath() + TEMPLET;
		return path;
	}
	/**
	 * 打印模版位置绝对路径
	 */
	public static String getPrintTempletPath() {
		String path = getRealServerPath() + PRINTTEMPLET;
		return path;
	}
	/**
	 * 采集表模版位置绝对路径
	 */
	public static String getTempletDate() {
		String path = getRealServerPath() + TEMPLETDATE;
		return path;
	}
	/**
	 * 街道数据
	 */
	public static String getDataTownStreet() {
		String path = getRealServerPath() + DATATOWNSTREET;
		return path;
	}

	/**
	 * 要下载文档绝对路径
	 */
	public static String getGeneDownload() {
		String path = getRealServerPath() + GENEDOWNLOAD;
		return path;
	}

	/**
	 *  生成文档位置（汇总）
	 */
	public static String getWRITEPATH() {
		String path = getRealServerPath() + WRITEPATH;
		return path;
	}

	public static void getAllMap() {
		// 获得工作簿的引用
		Workbook workbook = Constants.getWorkBook(Constants.getTempletPath());
//		Workbook workbookp = Constants.getWorkBook(Constants.getPrintTempletPath());
		Workbook dataStreet = Constants.getWorkBook(Constants.getDataTownStreet());
		if (workbook instanceof HSSFWorkbook) {
		} else if (workbook instanceof XSSFWorkbook) {
			XSSFRead(workbook);
//			ReadXSSF4Con(workbook);
//			XSSFReadPrint(workbookp);
			XSSFReadStreetData(dataStreet);
			try {
				workbook.close();
				dataStreet.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
	/**
	 * 读取模版表头信息（字段所在列）
	 * @param simWorkbook
	 */
	public static void XSSFRead(Workbook simWorkbook) {
		XSSFWorkbook workbook = (XSSFWorkbook) simWorkbook;
		// 创建对工作表的引用。
		XSSFSheet sheet = workbook.getSheetAt(1);
		for (int r = 2; r < 5; r++) {
			XSSFRow row = sheet.getRow(r);
			int cellNum = row.getLastCellNum();

			for (int c = 1; c < cellNum; c++) {
				XSSFCell cell = row.getCell(c);
				if (null == cell || null == cell.getStringCellValue())
					continue;
				String cellName = cell.getStringCellValue().trim()
						.replace(" ", "").replaceAll("\\n", "")
						.replaceAll("：", ":").split("\\(")[0].split("\\（")[0];
				if("党员人数".equals(cellName)){
					totalMap.put("专职党员", c + "");
					totalMapRec.put(c + "", "专职党员");
					continue;
				}
				if("党员数".equals(cellName)){
					totalMap.put("驻厦单位党员情况党员人数", c + "");
					totalMapRec.put(c + "", "驻厦单位党员情况党员人数");
					continue;
				}
				cellName = cellName.contains(":") ? cellName.split(":")[1]
						: cellName;
				cellName = cellName.contains("数") ? cellName.split("数")[0]
						: cellName;
				cellName = cellName.endsWith("人") ? cellName.substring(0,
						cellName.length() - 1) : cellName;
				cellName = cellName.replace("的", "");
				cellName = cellName.replace("楼宇", "");
				cellName = cellName.replace("工作", "");
				if(cellName.contains("建立")&&(!cellName.contains("是否"))){
					cellName = cellName.substring(cellName.lastIndexOf("建立")+2,cellName.length());
				}
				totalMap.put(cellName, c + "");
				totalMapRec.put(c + "", cellName);
			}
		}
		
		XSSFRow row8 = sheet.getRow(7);//累计 行
		int cindex =0;
		String cFor = "";
		for(Cell c : row8 ){
			cindex = c.getColumnIndex();
			cFor = c.getCellType()==Cell.CELL_TYPE_FORMULA?c.getCellFormula():c.getStringCellValue();
			forMap.put(cindex, cFor);
			forMapRec.put( cFor,cindex);
		}
		
	}
	/**
	 * 读取打印模版
	 * @param simWorkbook
	 */
	public static void XSSFReadPrint(Workbook simWorkbook) {
		XSSFWorkbook workbook = (XSSFWorkbook) simWorkbook;
		// 创建对工作表的引用。
		// 本例是按名引用（让我们假定那张表有着缺省名"Sheet1"）
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			XSSFSheet sheet = workbook.getSheetAt(i);
			for (int r = 2; r < 5; r++) {
				XSSFRow row = sheet.getRow(r);
				int cellNum = row.getLastCellNum();
				for (int c = 1; c < cellNum; c++) {
					XSSFCell cell = row.getCell(c);
					if (null == cell || null == cell.getStringCellValue())
						continue;
					String cellName = cell.getStringCellValue().trim()
							.replace(" ", "").replaceAll("\\n", "")
							.replaceAll("：", ":").split("\\(")[0].split("\\（")[0];
					if("党员人数".equals(cellName)){
						p1Map.put("专职党员", c + "");
						p1MapRec.put(c + "", "专职党员");
						continue;
					}
					if("党员数".equals(cellName)){
						p1Map.put("驻厦单位党员情况党员人数", c + "");
						p1MapRec.put(c + "", "驻厦单位党员情况党员人数");
						continue;
					}
					cellName = cellName.contains(":") ? cellName.split(":")[1]
							: cellName;
					cellName = cellName.contains("数") ? cellName.split("数")[0]
							: cellName;
					cellName = cellName.endsWith("人") ? cellName.substring(0,
							cellName.length() - 1) : cellName;
					cellName = cellName.replace("的", "");
					cellName = cellName.replace("楼宇", "");
					cellName = cellName.replace("工作", "");
					if (cellName.contains("建立") && (!cellName.contains("是否"))) {
						cellName = cellName.substring(
								cellName.lastIndexOf("建立") + 2,
								cellName.length());
					}
					switch(i){
					case 0:
						p1Map.put(cellName, c + "");
						p1MapRec.put(c + "", cellName);
						break;
					case 1:
						p2Map.put(cellName, c + "");
						p2MapRec.put(c + "", cellName);
						break;
					case 2:
						p3Map.put(cellName, c + "");
						p3MapRec.put(c + "", cellName);
						break;
					case 3:
						p4Map.put(cellName, c + "");
						p4MapRec.put(c + "", cellName);
						break;
					}
					
				}
			}

		}
	}
	/**
	 * 读取街道数据
	 * @param simWorkbook
	 */
	public static void XSSFReadStreetData(Workbook simWorkbook) {
		Sheet dataSheet = simWorkbook.getSheetAt(1);
	
		
		String lev ="";
		String code = "";
		String cap = "";
		String pcode = "";
		for(Row r : dataSheet){
			lev = r.getCell(4).getStringCellValue();
			code = r.getCell(0).getStringCellValue();
			pcode = r.getCell(3).getStringCellValue();
			if("3".equals(lev)){
				cap = r.getCell(1).getStringCellValue();
				allTown.put(code, cap);
				Map<String,String> tmp = allStreet.get(code);
				if(null == tmp){
					allStreet.put(code,  new HashMap<String,String>());
				}
			}else if("4".equals(lev)){
				Map<String,String> tmp = allStreet.get(pcode);
				cap = r.getCell(2).getStringCellValue();
				if(null == tmp){
					tmp = new HashMap<String,String>();
					tmp.put(code, cap);
					allStreet.put(pcode, tmp);
				}else{
					tmp.put(code, cap);
					allStreet.put(pcode, tmp);
				}
			}
		}
		
	}
	//最早应该是读取采集表 字段名及位置 已废弃 且有问题 索引不对
	/*public static void ReadXSSF4Con(Workbook simWorkbook) {
		XSSFWorkbook workbook = (XSSFWorkbook) simWorkbook;
		// 创建对工作表的引用。
		// 本例是按名引用（让我们假定那张表有着缺省名"Sheet1"）
		XSSFSheet sheet = workbook.getSheetAt(1);
		int rowNum = sheet.getLastRowNum();
		for (int r = 0; r < rowNum; r++) {
			XSSFRow row = sheet.getRow(r);
			int cellNum = row.getLastCellNum();
			for (int c = 0; c < cellNum; c++) {
				XSSFCell cell = row.getCell(c);
				if (null == cell) {
					continue;
				} else {
					int t = cell.getCellType();
					String value = "";
					switch (t) {
					case 0:
						value = cell.getNumericCellValue() + "";
						break;
					case 1:
						value = cell.getStringCellValue();
						break;
					case 2:
						value = cell.getCellFormula();
						break;
					case 3:
						value = "";
						break;
					case 4:
						value = cell.getBooleanCellValue() + "";
						break;
					case 5:
						value = cell.getErrorCellString();
						break;
					}
					String cellName = value.trim().replace(" ", "")
							.replaceAll("\\n", "").replaceAll("：", ":")
							.split("\\(")[0].split("\\（")[0];
					cellName = cellName.contains(":") ? cellName.split(":")[1]
							: cellName;
					cellName = cellName.contains("数") ? cellName.split("数")[0]
							: cellName;
					cellName = cellName.endsWith("人") ? cellName.substring(0,
							cellName.length() - 1) : cellName;
					cellName = cellName.replace("的", "");
					cellName = cellName.replace("楼宇", "");
					cellName = cellName.replace("工作", "");
					if(cellName.contains("建立")&&(!cellName.contains("是否"))){
						cellName = cellName.substring(cellName.lastIndexOf("建立")+2,cellName.length());
					}
//					basicMap.put(cellName, r + "." + c);
//					basicMapRec.put(r + "." + c, cellName);
				}
			}
		}
	}
*/
	public static Workbook getWorkBook(String path) {
		Workbook workbook = null;
		try {
			FileInputStream fis = new FileInputStream(path);
			workbook = new HSSFWorkbook(fis);
			fis.close();
		} catch (OfficeXmlFileException e) {
			// TODO Auto-generated catch block

			try {
				FileInputStream fis = new FileInputStream(path);
				workbook = new XSSFWorkbook(fis);
				fis.close();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				return workbook;
			} 
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return workbook;
	}
}
