package com.lydj.Controller.util;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.Map;
import java.util.Random;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class utils {
	/**
	 * 生产随机字符串
	 * @param length
	 * @return
	 */
	public static String getRandomString(){
		String s =geneRandomString(Constants.RONDOMLENGTH);
		return s.toString();
		
	}
	public static String geneRandomString(int length){
	    Random random=new Random();
	    StringBuffer sb=new StringBuffer();
	    for(int i=0;i<length;i++){
	       int number=random.nextInt(3);
	       long result=0;
	       switch(number){
	          case 0:
	              result=Math.round(Math.random()*25+65);
	              sb.append(String.valueOf((char)result));
	              break;
	         case 1:
	             result=Math.round(Math.random()*25+97);
	             sb.append(String.valueOf((char)result));
	             break;
	         case 2:     
	             sb.append(String.valueOf(new Random().nextInt(10)));
	             break;
	        }
	   
	     }
	     return sb.toString();
	 }
	
	
	/**
	 * zip文件解压缩
	 * @return
	 * @throws FileNotFoundException 
	 * @throws IOException 
	 */
	public static boolean zipDeco(File f) {
		long startTime=System.currentTimeMillis();  
		if(f.isDirectory()){
			File[] fs = f.listFiles();
			for(File file :fs){
				
				String Parent=file.getParentFile().getPath(); //输出路径（文件夹目录）  
				File Fout=null;  
				ZipEntry entry;  
				try {
					ZipInputStream Zin=new ZipInputStream(new FileInputStream(file),Charset.forName("GBK"));//输入源zip路径  
					BufferedInputStream Bin=new BufferedInputStream(Zin);  
					while((entry = Zin.getNextEntry())!=null && !entry.isDirectory()){  
						Fout=new File(Parent,entry.getName());  
						if(!Fout.exists()){  
							(new File(Fout.getParent())).mkdirs();  
						}  
						FileOutputStream out=new FileOutputStream(Fout);  
						BufferedOutputStream Bout=new BufferedOutputStream(out);  
						int b;  
						while((b=Bin.read())!=-1){  
							Bout.write(b);  
						}  
						Bout.close();  
						out.close();  
						System.out.println(Fout+"解压成功");      
					}  
					Bin.close();  
					Zin.close();
				} catch (FileNotFoundException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
					return false;
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
					return false;
				}  
			}
		}
    long endTime=System.currentTimeMillis();  
    System.out.println("耗费时间： "+(endTime-startTime)+" ms");  
		
		return true;
	}
	
	public static void setCellVall(Cell cell, String value){
		 Pattern pattern = Pattern.compile("-?\\d+(\\.)?(\\d+)?");
		 if(null == value) {
			 cell.setCellValue("");
		 }else{
			 Matcher isNum = pattern.matcher(value);
			if( !isNum.matches() ){
//		        str 
				 cell.setCellValue(value);
			 }else{
				 // int
				 DecimalFormat df = new DecimalFormat("#");
				 cell.setCellValue(Double.parseDouble(df.format(Double.parseDouble(value))));
			 }
		 }
	}
	
	/**
	 * 合并Excel
	 * 
	 * @param toFileName
	 *            写入的文件路径
	 * @param toSheetIndex    
	 *    		     写入的文件SheetIndex，如果写在最后请设置-1，否则请在Sheet数量范围内
	 * @param fromFileName
	 *            读取的文件路径
	 * @param fromSheetIndex
	 *            读取的文件SheetIndex
	 * @throws Exception
	 */
	public static void mergeExcel(String toFileName, int toSheetIndex, String fromFileName,
			int fromSheetIndex) throws Exception {
			
			// 1、打开Excel1
			/*InputStream inputStream = new FileInputStream(toFileName);
			XSSFWorkbook toWorkbook = new XSSFWorkbook(inputStream);
			inputStream.close();

			// 2、打开Excel2
			inputStream = new FileInputStream(fromFileName);
			XSSFWorkbook fromWorkbook = new XSSFWorkbook(inputStream);
			inputStream.close();*/

			// 3、复制Sheet，放在ToExcel1的Sheet上
//			long start = System.currentTimeMillis();
//			for(int i =0;i<1000;i++){
//				copySheet(toWorkbook, toSheetIndex, fromWorkbook, fromSheetIndex);
//			}
//			long stop = System.currentTimeMillis();
//			System.out.println("copySheet："+(stop - start));
//			
//			// 写入Excel1文件
//			OutputStream outputStream = new FileOutputStream(toFileName);
//			toWorkbook.write(outputStream);
//			outputStream.close();
	}
/**
 * 
 * @param toWorkbook
 * @param toSheetIndex
 * @param fromWorkbook
 * @param fromSheetIndex
 * @throws Exception
 */
public  static void copySheet(XSSFWorkbook toWorkbook, int toSheetIndex,XSSFWorkbook fromWorkbook, int fromSheetIndex,String sheetName) throws Exception {
//	Sheet fromSheet = fromWorkbook.cloneSheet(fromSheetIndex); 
	Sheet fromSheet = fromWorkbook.getSheetAt(fromSheetIndex);
	
//	String sheetName = fromSheet.getSheetName().replace("(2)", "")
//			+ "无样式"+utils.geneRandomString(5);
	Sheet toSheet = toWorkbook.getSheet(sheetName);
	if (null == toSheet) {
		toSheet = toWorkbook.createSheet(sheetName);
		if (toSheetIndex >= 0) {
			toWorkbook.setSheetOrder(sheetName, toSheetIndex);
		}
	} else {
		throw new Exception("相同名称的Sheet已存在");
	}
	toSheet.setDisplayGridlines(true);
	// 1、合并单元格
	for (int mrIndex = 0; mrIndex < fromSheet.getNumMergedRegions(); mrIndex++) {
		CellRangeAddress cellRangeAddress = fromSheet
				.getMergedRegion(mrIndex);
		toSheet.addMergedRegion(cellRangeAddress);
	}

		// 2、单元格赋值，样式等
		Map<Integer, Integer> setColumnWidthIndex = new HashMap<Integer, Integer>();
		for (int rIndex = fromSheet.getFirstRowNum(); rIndex <= fromSheet
				.getLastRowNum(); rIndex++) {
			Row fromRow = fromSheet.getRow(rIndex);
			if (null == fromRow) {
				continue;
			}
			Row toRow = toSheet.createRow(rIndex);

			// 设置行高，自动行高即可
			toRow.setHeight(fromRow.getHeight());
			
			XSSFCellStyle toCellStyle = toWorkbook.createCellStyle();
			toCellStyle.setBorderBottom(CellStyle.BORDER_THIN);//下边框       
			toCellStyle.setBorderLeft(CellStyle.BORDER_THIN);//左边框       
			toCellStyle.setBorderRight(CellStyle.BORDER_THIN);//右边框       
			toCellStyle.setBorderTop(CellStyle.BORDER_THIN);//上边框    
			// 设置Cell的值和样式
			for (int cIndex = fromRow.getFirstCellNum(); cIndex <= fromRow
					.getLastCellNum(); cIndex++) {
				Cell fromCell = fromRow.getCell(cIndex);
				if (null == fromCell) {
					continue;
				}
				Cell toCell = toRow.createCell(cIndex);
				toCell.setCellStyle(toCellStyle);
				// 设置列宽
				Integer isSet = setColumnWidthIndex.get(cIndex);
				if (null == isSet) {
					toSheet.setColumnWidth(cIndex,
							fromSheet.getColumnWidth(cIndex));
					setColumnWidthIndex.put(cIndex, cIndex);
				}
			/*// 设置单元格样式
			CellStyle fromCellStyle = fromCell.getCellStyle();
			if (fromCellStyle.getIndex() != 0) {
//				XSSFCellStyle toCellStyle = toWorkbook.getCellStyleAt((short)0);
				XSSFCellStyle toCellStyle = toWorkbook.createCellStyle();
//				XSSFCellStyle toCellStyle = (XSSFCellStyle) fromCellStyle;
				// 文字展示样式
				String fromDataFormat = fromCellStyle.getDataFormatString();
				if ((null != fromDataFormat)
						&& ("".equals(fromDataFormat) == false)) {
					DataFormat toDataFormat = toWorkbook.createDataFormat();
					toCellStyle.setDataFormat(toDataFormat
							.getFormat(fromDataFormat));

				}
				// 文字换行
				toCellStyle.setWrapText(fromCellStyle.getWrapText());
				// 文字对齐方式
				toCellStyle.setAlignment(fromCellStyle.getAlignment());
				toCellStyle.setVerticalAlignment(fromCellStyle
						.getVerticalAlignment());
				// 单元格边框
				toCellStyle.setBorderLeft(fromCellStyle.getBorderLeft());
				toCellStyle.setBorderRight(fromCellStyle.getBorderRight());
				toCellStyle.setBorderTop(fromCellStyle.getBorderTop());
				toCellStyle
						.setBorderBottom(fromCellStyle.getBorderBottom());
				// 字体颜色，大小
				short fromFontIndex = fromCellStyle.getFontIndex();
				XSSFFont fromFont = fromWorkbook.getFontAt(fromFontIndex);
				Short toFontIndex = setFontIndex.get(fromFontIndex);
				if (null == toFontIndex) {
					XSSFFont toFont = toWorkbook.getFontAt(fromFontIndex);
					
					toFont.setBoldweight(fromFont.getBoldweight());
					  toFont.setFontName(fromFont.getFontName());
					 toFont.setFontHeightInPoints(fromFont
					 .getFontHeightInPoints());
					 toFont.setColor(fromFont.getXSSFColor());
					 toFont.setBold(fromFont.getBold());
					 
					toCellStyle.setFont(toFont);
					// 设置的Font加入集合中
					toFontIndex = toFont.getIndex();
					setFontIndex.put(fromFontIndex, toFontIndex);
				} else {
					XSSFFont toFont = toWorkbook.getFontAt(toFontIndex);
					toCellStyle.setFont(toFont);
				}
			
				// 背景色
				XSSFColor fillForegroundColor = (XSSFColor) fromCellStyle
						.getFillForegroundColorColor();
				toCellStyle.setFillForegroundColor(fillForegroundColor);
				toCellStyle.setFillPattern(fromCellStyle.getFillPattern());

				toCell.setCellStyle(toCellStyle);
			}*/
			int fromCellType = fromCell.getCellType();
			switch (fromCellType) {
			case Cell.CELL_TYPE_STRING:
				toCell.setCellValue(fromCell.getStringCellValue());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				toCell.setCellValue(fromCell.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				toCell.setCellValue(fromCell.getBooleanCellValue());
				break;
			default:
				toCell.setCellValue(fromCell.getStringCellValue());
				break;
			}
		}
	}
}

/**
 * 
 * @param toWorkbook
 * @param toSheetIndex  已无效果
 * @param fromWorkbook
 * @param fromSheetIndex
 * @throws Exception
 */
public  static void copySheetTest(XSSFWorkbook toWorkbook, int toSheetIndex,XSSFWorkbook fromWorkbook,
		int fromSheetIndex,String sheetName,XSSFCellStyle toCellStyle,XSSFCellStyle headStyle ,boolean copyStyleFlag) throws Exception {
//	Sheet fromSheet = fromWorkbook.cloneSheet(fromSheetIndex); 
	Sheet fromSheet = fromWorkbook.getSheetAt(fromSheetIndex);
	Sheet toSheet = toWorkbook.getSheet(sheetName);
	if (null == toSheet) {
		toSheet = toWorkbook.createSheet(sheetName);
//		if (toSheetIndex >= 0) {
			toWorkbook.setSheetOrder(sheetName, toWorkbook.getNumberOfSheets()-1);
//		}
	} else {
		throw new Exception("相同名称的Sheet已存在");
	}
	toSheet.setDisplayGridlines(true);
	// 1、合并单元格
	CellRangeAddress cellRangeAddress;
	for (int mrIndex = 0; mrIndex < fromSheet.getNumMergedRegions(); mrIndex++) {
		 cellRangeAddress = fromSheet
				.getMergedRegion(mrIndex);
		toSheet.addMergedRegion(cellRangeAddress);
	}

		// 2、单元格赋值，样式等
		Map<Integer, Integer> setColumnWidthIndex = new HashMap<Integer, Integer>();
		Map<Short, Short> setFontIndex = new HashMap<Short, Short>();
		Row fromRow ;
		Row toRow ;
		Cell fromCell;
		Cell toCell;
		Integer isSet;
		int fromCellType;
		for (int rIndex = fromSheet.getFirstRowNum(); rIndex <= fromSheet.getLastRowNum(); rIndex++) {
			fromRow = fromSheet.getRow(rIndex);
			if (null == fromRow) {
				continue;
			}
			toRow = toSheet.createRow(rIndex);

			// 设置行高，自动行高即可
			toRow.setHeight(fromRow.getHeight());
			
			
//			XSSFCellStyle toCellStyle = toWorkbook.createCellStyle();
//			toCellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);//下边框       
//			toCellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框       
//			toCellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框       
//			toCellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框    
			// 设置Cell的值和样式
			for (int cIndex = fromRow.getFirstCellNum(); (cIndex != -1)
					&& (cIndex <= fromRow.getLastCellNum()); cIndex++) {
				fromCell = fromRow.getCell(cIndex);
				if (null == fromCell) {
					continue;
				}
				toCell = toRow.createCell(cIndex);
				toCell.setCellStyle(toCellStyle);
				// 设置列宽
				isSet = setColumnWidthIndex.get(cIndex);
				if (null == isSet) {
					toSheet.setColumnWidth(cIndex,
							fromSheet.getColumnWidth(cIndex));
					setColumnWidthIndex.put(cIndex, cIndex);
				}
				// 设置单元格样式
				if (copyStyleFlag) {
					CellStyle fromCellStyle = fromCell.getCellStyle();
					if (fromCellStyle.getIndex() != 0) {
						// XSSFCellStyle toCellStyle =
						// toWorkbook.getCellStyleAt((short)0);
						toCellStyle = toWorkbook.createCellStyle();
						// XSSFCellStyle toCellStyle = (XSSFCellStyle)
						// fromCellStyle;
						// 文字展示样式
						String fromDataFormat = fromCellStyle
								.getDataFormatString();
						if ((null != fromDataFormat)
								&& ("".equals(fromDataFormat) == false)) {
							DataFormat toDataFormat = toWorkbook
									.createDataFormat();
							toCellStyle.setDataFormat(toDataFormat
									.getFormat(fromDataFormat));

						}
						// 文字换行
						toCellStyle.setWrapText(fromCellStyle.getWrapText());
						// 文字对齐方式
						toCellStyle.setAlignment(fromCellStyle.getAlignment());
						toCellStyle.setVerticalAlignment(fromCellStyle
								.getVerticalAlignment());
						// 单元格边框
						toCellStyle
								.setBorderLeft(fromCellStyle.getBorderLeft());
						toCellStyle.setBorderRight(fromCellStyle
								.getBorderRight());
						toCellStyle.setBorderTop(fromCellStyle.getBorderTop());
						toCellStyle.setBorderBottom(fromCellStyle
								.getBorderBottom());
						// 字体颜色，大小
						short fromFontIndex = fromCellStyle.getFontIndex();
						XSSFFont fromFont = fromWorkbook
								.getFontAt(fromFontIndex);
						Short toFontIndex = setFontIndex.get(fromFontIndex);
						if (null == toFontIndex) {
							XSSFFont toFont = toWorkbook
									.getFontAt(fromFontIndex);

							toFont.setBoldweight(fromFont.getBoldweight());
							toFont.setFontName(fromFont.getFontName());
							toFont.setFontHeightInPoints(fromFont
									.getFontHeightInPoints());
							toFont.setColor(fromFont.getXSSFColor());
							toFont.setBold(fromFont.getBold());

							toCellStyle.setFont(toFont);
							// 设置的Font加入集合中
							toFontIndex = toFont.getIndex();
							setFontIndex.put(fromFontIndex, toFontIndex);
						} else {
							XSSFFont toFont = toWorkbook.getFontAt(toFontIndex);
							toCellStyle.setFont(toFont);
						}

						// 背景色
						// XSSFColor fillForegroundColor = (XSSFColor)
						// fromCellStyle
						// .getFillForegroundColorColor();
						// toCellStyle.setFillForegroundColor(fillForegroundColor);
						// toCellStyle.setFillPattern(fromCellStyle.getFillPattern());

						toCell.setCellStyle(toCellStyle);
					}
				}
				fromCellType = fromCell.getCellType();
				switch (fromCellType) {
				case Cell.CELL_TYPE_STRING:
					toCell.setCellValue(fromCell.getStringCellValue());
					break;
				case Cell.CELL_TYPE_NUMERIC:
					toCell.setCellValue(fromCell.getNumericCellValue());
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					toCell.setCellValue(fromCell.getBooleanCellValue());
					break;
				case Cell.CELL_TYPE_FORMULA:
					toCell.setCellFormula(fromCell.getCellFormula());
					break;
				default:
					toCell.setCellValue(fromCell.getStringCellValue());
					break;
				}
			}
	}
		toSheet.getRow(0).getCell(0).setCellStyle(headStyle);
//		toWorkbook.close();
}


}
