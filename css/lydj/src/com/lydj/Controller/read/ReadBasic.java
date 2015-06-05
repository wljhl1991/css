package com.lydj.Controller.read;

import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 读取采集表
 * @author ldj
 *
 */
public class ReadBasic {

	/*public static Map<String, String> getbasic(String[] args) {
		Map<String,String> map = null ;
		// 获得工作簿的引用
		Workbook workbook = ReadXL.getWorkBook(Constants.getTestRead());
		if(	workbook instanceof HSSFWorkbook){
		}else if(	workbook instanceof XSSFWorkbook){
			 map = XSSFRead(workbook,2);
		}
		return map;
	}*/
	
	
	public static Map <String,String>   XSSFRead(Workbook simWorkbook,int sheetIndex){
		Map <String,String> basic = new HashMap<String,String>();
		try{
			if(	simWorkbook instanceof HSSFWorkbook){
				
			}else if(	simWorkbook instanceof XSSFWorkbook){
				XSSFWorkbook workbook  = (XSSFWorkbook)simWorkbook;
				
				DecimalFormat df = new DecimalFormat("0.0");//保留两位小数且不用科学计数法
				df.setRoundingMode(RoundingMode.HALF_UP); //四舍五入
				// 创建对工作表的引用。
				// 本例是按名引用（让我们假定那张表有着缺省名"Sheet1"）
				XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
				//最大行数
				int rowNum = sheet.getLastRowNum();
				//最大列数
				int cellNum = 0;
				//记录 专职工作人员人数 所在位置 取得其中党员数 50 为默认行数
				int zzry =50;
				String key  = "";
				String value = "";
				for(int i =0;i<rowNum;i++){
					XSSFRow row = sheet.getRow(i);
					if(null== row) continue;
					cellNum = row.getLastCellNum();
					for(int j = 0;j<cellNum;j++){
						XSSFCell cell = row.getCell(j);
						if( null == cell ){
							System.out.println(row+","+j);
							continue;	
						}
						int t = cell.getCellType();
						switch(t){
//				CELL_TYPE_NUMERIC 数值型 0
//				CELL_TYPE_STRING 字符串型 1
//				CELL_TYPE_FORMULA 公式型 2
//				CELL_TYPE_BLANK 空值 3
//				CELL_TYPE_BOOLEAN 布尔型 4
//				CELL_TYPE_ERROR 错误 5
						case 0: value = df.format(new Double(cell.getNumericCellValue()))+"";
						value = value.endsWith(".0")?value.substring(0,value.lastIndexOf(".")):value;
						break;
						case 1:	value = cell.getStringCellValue().trim().replace(" ", "").replaceAll("\\n", "");
						break;
						case 2: value = cell.getCellFormula();break;
						case 3: value = "";break;
						case 4: value = cell.getBooleanCellValue()+"";break;
						case 5: value = cell.getErrorCellString();break;
						}
						if(j%2==0){
							if("党员人数".equals(value)){
//								key = value;
								key = "驻厦单位党员情况党员人数";
								continue;
							}
							if("专职工作人员人数".equals(value)){
//								key = value;
								zzry = i;
							}
							value = value.replaceAll("：", ":").split("\\(")[0].split("\\（")[0];
							if(value.indexOf(":") != value.length()-1){
								value = value.contains(":")?value.split(":")[1]:value;
							}
							value = value.contains("数")?value.split("数")[0]:value;
							value = value.endsWith("人")?value.substring(0,value.length()-1):value;
							value = value.replace("的", "");
							value = value .replace("楼宇", "");
							value = value.replace("工作", "");
							if(value.contains("建立")&&(!value.contains("是否"))){
								value = value.substring(value.lastIndexOf("建立")+2,value.length());
							}
							/*if("党员人数".equals(value)){
								key = value;
								continue;
							}*/
							if("联合党组织".equals(value)){
								value =  row.getCell(2).getCellType() == Cell.CELL_TYPE_NUMERIC?row.getCell(2).getNumericCellValue()+"":row.getCell(2).getStringCellValue();
								value = value.endsWith(".0")?value.substring(0,value.lastIndexOf(".")):value;
								value = (null == value || "".equals(value) || "无".equals(value))?"0":value;
								basic.put("联合党组织",value);
								continue;
							}
							if("上一年度经费".equals(value)){
								value =  "2014年经费";
							}
							if("2014年新发展预备党员".equals(value)){
								value =  "2014年新发展党员";
							}
							key = value;
						}else{
							value = value.replaceAll("愿意亮明身份且能够证明", "");
							value = (null == value || "".equals(value) || "无".equals(value) )?"0":value;
							basic.put(key,value);
						}
					}
				}
				
				/*if(null != sheet.getRow(sheet.getLastRowNum()).getCell(1) && sheet.getRow(sheet.getLastRowNum()).getCell(1).getCellType() == XSSFCell.CELL_TYPE_NUMERIC){
					value = df.format(sheet.getRow(sheet.getLastRowNum()).getCell(1).getNumericCellValue())+"";
					value = value.endsWith(".0")?value.substring(0,value.lastIndexOf(".")):value;
					basic.put("其他途径",value);
				}else if(null != sheet.getRow(sheet.getLastRowNum()).getCell(1) && sheet.getRow(sheet.getLastRowNum()).getCell(1).getCellType() == XSSFCell.CELL_TYPE_STRING){
					value = sheet.getRow(sheet.getLastRowNum()).getCell(1).getStringCellValue()+"";
					value = value.endsWith(".0")?value.substring(0,value.lastIndexOf(".")):value;
					basic.put("其他途径",value);
				}*/
				
				
				value =  sheet.getRow(zzry+1).getCell(1).getCellType() == Cell.CELL_TYPE_NUMERIC?sheet.getRow(zzry+1).getCell(1).getNumericCellValue()+"":sheet.getRow(zzry+1).getCell(1).getStringCellValue();
				value = value.endsWith(".0")?value.substring(0,value.lastIndexOf(".")):value;
				value = (null == value || "".equals(value)|| "无".equals(value))?"0":value;
				basic.put("专职党员",value);
				workbook.close();
			}
			simWorkbook.close();
			
		}catch(Exception e){
			e.printStackTrace();
			return null;
		}
		return basic;
	}
}
