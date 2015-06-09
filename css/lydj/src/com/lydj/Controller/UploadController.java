package com.lydj.Controller;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;
import org.springframework.web.multipart.commons.CommonsMultipartResolver;
import org.springframework.web.servlet.ModelAndView;

import com.lydj.Controller.read.ReadBasic;
import com.lydj.Controller.read.ReadXL;
import com.lydj.Controller.read.WriteXL;
import com.lydj.Controller.util.ConInfo;
import com.lydj.Controller.util.Constants;
import com.lydj.Controller.util.utils;

@Controller
public class UploadController {
	 private static Logger logger = LoggerFactory.getLogger(UploadController.class);
	 
	 
	@RequestMapping("/upload2.do")
	public void upload2(HttpServletRequest request,HttpServletResponse response) throws IllegalStateException, IOException{
		String localForld = Constants.getRealServerPath()+"upload"+File.separator+request.getParameter("localFolad");
		CommonsMultipartResolver multipartResolver  = new CommonsMultipartResolver(request.getSession().getServletContext());
		if(multipartResolver.isMultipart(request)){
			MultipartHttpServletRequest  multiRequest = (MultipartHttpServletRequest)request;
			
			Iterator<String>  iter = multiRequest.getFileNames();
			while(iter.hasNext()){
					MultipartFile file = multiRequest.getFile(iter.next());
				if(file != null){
					if(!new File(localForld).exists()){
						new File(localForld).mkdir();
					}
					String fileName = file.getOriginalFilename();
					String path = localForld+ File.separator + System.currentTimeMillis() + fileName;
					
					File localFile = new File(path);
					
					file.transferTo(localFile);
				}
				
			}
		}
//		request.setAttribute("fileName", "fileName1");
		/*return "/success";*/
	}
	
	@RequestMapping("/toUpload.do")
	public String toUpload(HttpServletRequest request,HttpServletResponse response){
		String s = utils.getRandomString();
		//生产随机字符串 标记本次上传
		request.setAttribute("localFolad", s);
		request.setAttribute("fileName", "fileName1");
		request.setAttribute("town", ConInfo.TOWN);
		request.setAttribute("street", ConInfo.getStreet("110105000"));
		logger.debug("toUpload.do : localFolad :" +s);
		return "/indext";	
	}
	
	/**
	 * 生产环境测试使用
	 * @param request
	 * @param response
	 * @return
	 */
	@RequestMapping("/toUploadTest.do")
	public String toUploadTest(HttpServletRequest request,HttpServletResponse response){
		String s = utils.getRandomString();
		//生产随机字符串 标记本次上传
		request.setAttribute("localFolad", s);
		request.setAttribute("fileName", "fileName1");
		request.setAttribute("town", ConInfo.TOWN);
		request.setAttribute("street", ConInfo.getStreet("110105000"));
		logger.debug("toUploadTest.do : localFolad :" +s);
		return "/indexttll";	
	}
	 @ResponseBody
	@RequestMapping("/changeData.do")
	public Map changeData(HttpServletRequest request,HttpServletResponse response){
	
		Map m = ConInfo.getStreet(request.getParameter("id"));	
		return  m;
	}

	@ResponseBody 
	@RequestMapping("/start.do")
	public void start(HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		response.setContentType("application/json");
		
		 DecimalFormat decimalFormat=new DecimalFormat("#");
		 
		logger.debug("进入 start.do");
		Workbook workbook = ReadXL.getWorkBook(Constants.getTempletPath());
		long start = System.currentTimeMillis();
		String localForld = Constants.getRealServerPath() + "upload"
				+ File.separator + request.getParameter("localFolad");

		// 测试环境 忽略工作站名
		String testok = request.getParameter("testok");
		String testStyle = request.getParameter("testStyle");
		// 所有上传的
		List<Workbook> uploadwbs = ReadXL.getWorkBookList(localForld);

		//表头名称
		String titleName = "";
		// 开始写入的位置
		int rowPos = 5;
		// 文件数量
		int fileNum = uploadwbs.size();
		// sheet 页数量
//		int sheetNum = 0;

		// 数据校验 错误提示信息
		String msg = "";

		// 若有校验错误 则不再继续汇总 且不生成汇总文件及打印 true为校验正确
		boolean flag = true;
		// 工作站名称
		List<String> stationNames = new ArrayList<String>();
		// 所有工作站名称
		StringBuilder sb = new StringBuilder();

		// 单独处理党委
		List<String> alldws = new ArrayList<String>();
		// 单独处理党员
//		List<String> alldys = new ArrayList<String>();

		// 返回前台的文件名
		String fileName = utils.getRandomString() + "";
		int[] t = new int[fileNum];// 单独处理的数量
		if (workbook instanceof HSSFWorkbook) {
			logger.debug("模版格式错误!");
		} else if (workbook instanceof XSSFWorkbook) {
			// 数字 汉字 汇总表
			Map<String, String> totalMapRec = Constants.totalMap;
			// 模版
			XSSFWorkbook wb = (XSSFWorkbook) workbook;
			
			// 第一格 样式 复制时用
			XSSFCellStyle fheadStyle = wb.getSheetAt(2).getRow(0).getCell(0)
					.getCellStyle();
			XSSFCellStyle toCellStyle = wb.createCellStyle();
			toCellStyle.setBorderBottom(CellStyle.BORDER_THIN);// 下边框
			toCellStyle.setBorderLeft(CellStyle.BORDER_THIN);// 左边框
			toCellStyle.setBorderRight(CellStyle.BORDER_THIN);// 右边框
			toCellStyle.setBorderTop(CellStyle.BORDER_THIN);// 上边框
			XSSFCellStyle headStyle = wb.createCellStyle();
			headStyle.setFont(fheadStyle.getFont());

			//行 列 样式
			XSSFRow row5  = wb.getSheetAt(1).getRow(5);
			XSSFCellStyle hrowStyle = row5.getRowStyle();//实际为行格式  .getCell(1).getCellStyle()
//			XSSFCellStyle rowStyle = row5.getCell(1).getCellStyle();//实际的列格式
			
			Map<Integer,XSSFCellStyle> cellStylesMap = new HashMap<Integer,XSSFCellStyle>();
			int cellIndex =0;
			for(Cell c :row5){
				XSSFCellStyle cs = (XSSFCellStyle) c.getCellStyle();
				cs.setWrapText(true);
				cs.setAlignment(CellStyle.ALIGN_CENTER);
				cellStylesMap.put(cellIndex++, cs);
			}
			
			//累计行 列 样式
			XSSFRow row7  = wb.getSheetAt(1).getRow(7);
			XSSFCellStyle row7Style = row7.getRowStyle();//实际为行格式  .getCell(1).getCellStyle()
//			XSSFCellStyle rowStyle = row5.getCell(1).getCellStyle();//实际的列格式
			
			Map<Integer,XSSFCellStyle> r7cellStylesMap = new HashMap<Integer,XSSFCellStyle>();
			int r7cellIndex =0;
			for(Cell c :row7){
				XSSFCellStyle cs = (XSSFCellStyle) c.getCellStyle();
//				cs.setWrapText(true);
				r7cellStylesMap.put(r7cellIndex++, cs);
			}
//			wb.getSheetAt(1).shiftRows(9, 10, -1,true,true);
			for (int i = 0; i < fileNum; i++) {
				Workbook uploadwb = uploadwbs.get(i);
				// 应为获得模版汇总页
				if (uploadwb instanceof XSSFWorkbook) {

					// 数据采集页
					XSSFWorkbook fromwb = (XSSFWorkbook) uploadwb;
					if (fromwb.getNumberOfSheets() > 1) {
						response.getWriter()
								.write("{\"err\":\"错误\",\"result\":\"汇总类型选择错误,请更换汇总类型!\"}");
						uploadwb.close();
						wb.close();
						workbook.close();
						logger.debug("{\"err\":\"错误\",\"result\":\"汇总类型选择错误,请更换汇总类型!\"}");
						return;
					}

					// 得到采集表数据 索引0
					Map<String, String> readBasic = ReadBasic.XSSFRead(
							uploadwb, 0);

					String tosheetName = readBasic.get("商务站名称");
					titleName = readBasic.get("所属街道");
					if ("checked".equals(testok))
						tosheetName += utils.geneRandomString(4);
					if (sb.lastIndexOf(tosheetName + "reco") != -1 ) {
						int x = sb.lastIndexOf(tosheetName + "reco");
						if(x == 0 || "reco".equals(sb.substring(x-4))){
							msg = "工作站 " + tosheetName + "已经存在! 请修改后重新上传!";
							response.getWriter().write("{\"err\":\"错误\",\"result\":\"" + msg + "\"}");
							uploadwb.close();
							wb.close();
							workbook.close();
							logger.debug("{\"err\":\"错误\",\"result\":\"" + msg + "\"}");
							return;
						}
					}
					sb = sb.append(tosheetName + "reco");
					tosheetName = tosheetName.replace("；", "");
					stationNames.add(tosheetName);

					// 进行数据校验

					// 5.AD列填报“党员数”小于等于AC列“专职工作人员”填报数
					Double dys = Double
							.parseDouble((null == readBasic.get("专职党员")
									|| "".equals(readBasic.get("专职党员")) ? "0"
									: readBasic.get("专职党员")));
					Double zzgzry = Double
							.parseDouble((null == readBasic.get("专职人员")
									|| "".equals(readBasic.get("专职人员")) ? "0"
									: readBasic.get("专职人员")));
					if (dys > zzgzry) {
						msg += "工作站： " + tosheetName + " 中 专职工作人员数(" + zzgzry
								+ ") 小于其中党员数(" + dys + ")\\n<br/>";
						flag = false;
					}

					// 6.AG列填报“党建工作指导员”数小于等于AF列填报“兼职工作人员”数
					Double djzdy = Double
							.parseDouble((null == readBasic.get("党建指导员")
									|| "".equals(readBasic.get("党建指导员")) ? "0"
									: readBasic.get("党建指导员")));
					Double jzry = Double
							.parseDouble((null == readBasic.get("兼职人员")
									|| "".equals(readBasic.get("兼职人员")) ? "0"
									: readBasic.get("兼职人员")));
					if (djzdy > jzry) {
						msg += "工作站： " + tosheetName + " 中 党建工作指导员数(" + djzdy
								+ ") 大于兼职工作人员数(" + jzry + ")\\n<br/>";
						flag = false;
					}

					// 7.AK列“上一年度工作经费”=AL列“市级拨付”+AM列“区县拨付”+AN列“街道（乡镇）拨付”+AO列“其他途径”
					BigDecimal sndgzjf =new BigDecimal((null == readBasic
							.get("2014年经费")
							|| "".equals(readBasic.get("2014年经费")) ? "0"
							: readBasic.get("2014年经费")));
					BigDecimal sjbf = new BigDecimal
							((null == readBasic.get("市级拨付")
									|| "".equals(readBasic.get("市级拨付")) ? "0"
									: readBasic.get("市级拨付")));
					BigDecimal qxbf = new BigDecimal
							((null == readBasic.get("区县拨付")
									|| "".equals(readBasic.get("区县拨付")) ? "0"
									: readBasic.get("区县拨付")));
					BigDecimal  jdbf = new BigDecimal 
							((null == readBasic.get("街道拨付")
							|| "".equals(readBasic.get("街道拨付")) ? "0"
							: readBasic.get("街道拨付")));
					BigDecimal  qttj = new BigDecimal 
							((null == readBasic.get("其他途径")
									|| "".equals(readBasic.get("其他途径")) ? "0"
									: readBasic.get("其他途径")));

					if (!(sndgzjf.doubleValue() == (sjbf.add(qxbf).add( jdbf).add (qttj)).doubleValue())) {
						msg += "工作站： " + tosheetName + " 中 2014年工作经费(" + decimalFormat.format(sndgzjf)
								+ ") 与总来源数(" + decimalFormat.format(sjbf.add(qxbf).add( jdbf).add (qttj))
								+ ")不相等\\n<br/>";
						flag = false;
					}
					// wb.setSheetName(4, tosheetName);

					// 校验正确时执行
					if (flag) {
						boolean copyStyleFlag  = Constants.SWITCHSTYLECOPY;
						if ("checked".equals(testStyle)) copyStyleFlag = true;
						utils.copySheetTest(wb, 2, fromwb, 0, tosheetName,
								toCellStyle, headStyle,copyStyleFlag);
						// 模版汇总页
						XSSFSheet datasheet = wb.getSheetAt(1);
						XSSFRow row6 = null;
						// if(i>6){
						row6 = datasheet.createRow(rowPos + i);
						row6.setRowStyle(hrowStyle);
						// }else{
						// row6 = datasheet.getRow(rowPos + i);
						// }
						Set<String> keys = totalMapRec.keySet();
						String value = "";// 采集表中的值
						for (String k : keys) {
							value = readBasic.get(k);
							int index = Integer.parseInt(totalMapRec.get(k));
							XSSFCell cell = row6.getCell(index);
							if(index > 54) continue;
							if (null == cell) {
								cell = row6.createCell(index);
								cell.setCellStyle(cellStylesMap.get(index));
								cell.setCellType(Cell.CELL_TYPE_NUMERIC);
							}
							if (k.startsWith("覆盖楼栋")) {
								value = ("".equals(value.trim())) ? "0" : value;
								t[i] = Integer.parseInt(value);
								continue;
							} else if ("其他".equals(k)) {
								value = readBasic.get("其他单位");
//								cell.setCellValue(value);
								utils.setCellVall(cell,value);

							} else if ("党委".equals(k)) {
								value = readBasic.get("党委");
								value = (null == value || "".equals(value))?"0":value;
								alldws.add(value);
							} else if ("党总支".equals(k)) {
								value = readBasic.get("总支部");
//								cell.setCellValue(value);
								utils.setCellVall(cell,value);

							} else if ("党支部".equals(k)) {
								value = readBasic.get("支部");
//								cell.setCellValue(value);
								utils.setCellVall(cell,value);

							} else if ("党费返还及党组织和活动经费".equals(k)) {
								value = readBasic.get("党费返还以及党组织和活动经费");
//								cell.setCellValue(value);
								utils.setCellVall(cell,value);

							} else if ("街道经费".equals(k)) {
								value = readBasic.get("街道拨付");
//								cell.setCellValue(value);
								utils.setCellVall(cell,value);

							} else if ("".equals(k)) {

							} else {
//								cell.setCellValue(value);
								utils.setCellVall(cell,value);
							}
						}
					}
				} else {
					uploadwb.close();
					wb.close(); 
					workbook.close();
					msg = "第"+(i+1)+"个文件格式错误，实际为xls文件!";
					response.getWriter().write("{\"err\":\"错误\",\"result\":\"" + msg + "\"}");
					logger.debug("{\"err\":\"错误\",\"result\":\"" + msg + "\"}");
					return;
				}
				uploadwb.close();
			}
			// 校验正确时执行
			if (flag) {
				
				wb.getSheetAt(0).getRow(4).getCell(0).setCellValue("2014年"+titleName+"商务楼宇工作站情况汇总表");
				wb.getSheetAt(1).getRow(0).getCell(0).setCellValue("2014年"+titleName+"商务楼宇工作站情况汇总表");
				
				// 工作站名连接到sheet页
				CreationHelper createHelper = workbook.getCreationHelper();

				CellStyle hlink_style = wb.createCellStyle();
				hlink_style.setBorderBottom(CellStyle.BORDER_THIN);// 下边框
				hlink_style.setBorderLeft(CellStyle.BORDER_THIN);// 左边框
				hlink_style.setBorderRight(CellStyle.BORDER_THIN);// 右边框
				hlink_style.setBorderTop(CellStyle.BORDER_THIN);// 上边框
				hlink_style.setAlignment(CellStyle.ALIGN_CENTER);
				hlink_style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
				hlink_style.setWrapText(true);
				Font hlink_font = wb.createFont();
//				hlink_font.setUnderline(Font.U_SINGLE);
				hlink_font.setColor(IndexedColors.BLUE.getIndex());
				hlink_style.setFont(hlink_font);

				// 特殊处理 ( 覆盖楼宇楼栋数、 超链接)
				for (int k = 0; k < fileNum; k++) {

					try {
						Hyperlink link = createHelper
								.createHyperlink(org.apache.poi.common.usermodel.Hyperlink.LINK_DOCUMENT);
						Row spcRow = wb.getSheetAt(1).getRow(k + rowPos);
						String sname = stationNames.get(k);// 工作站名称
						spcRow.createCell(0).setCellValue(k + 1);// 序号
						spcRow.getCell(0).setCellStyle(cellStylesMap.get(0));
						spcRow.getCell(4).setCellValue(t[k]);// 覆盖楼宇
						link.setAddress("" + sname + "!A1");
						spcRow.getCell(3).setHyperlink(link);
						spcRow.getCell(3).setCellStyle(hlink_style);
						spcRow.getCell(16).setCellType(Cell.CELL_TYPE_NUMERIC);
						spcRow.getCell(16).setCellValue(Integer.parseInt(alldws.get(k)));
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}

				}
				
				Row row8 = wb.getSheetAt(1).createRow(5+fileNum);//.createCell(4).setCellFormula("SUM(E6,E7)");
				row8.setRowStyle(row7Style);
				Map<Integer, String> m = Constants.forMap;
				Set<Integer> ks = m.keySet();
				CellRangeAddress cellRangeAddress = new CellRangeAddress(5+fileNum, 5+fileNum, 0, 3);  
			    wb.getSheetAt(1).addMergedRegion(cellRangeAddress);  
			    row8.createCell(0).setCellValue("累计");
				for(Integer k : ks){
					if(k == 0|| k == 1 || k == 2 ||k == 3){
						XSSFCellStyle tem = r7cellStylesMap.get(k);
						tem.setAlignment(CellStyle.ALIGN_CENTER);
					};
					Cell c8 =row8.createCell(k);
					c8.setCellStyle(r7cellStylesMap.get(k));
					String fs = m.get(k);
					if(fs.contains(":")){
						fs = fs.replace("7", (5+fileNum)+"");
						fs = fs.replace("AC8=0", "AC"+(6+fileNum)+"=0");
						fs = fs.replace("AF8=0", "AF"+(6+fileNum)+"=0");
						fs = fs.replace("AH8=0", "AH"+(6+fileNum)+"=0");
						fs = fs.replace("AC8,0", "AC"+(6+fileNum)+",0");
						fs = fs.replace("AF8,0", "AF"+(6+fileNum)+",0");
						fs = fs.replace("AH8,0", "AH"+(6+fileNum)+",0");
						c8.setCellType(Cell.CELL_TYPE_FORMULA);
						c8.setCellFormula(fs);
					}
					c8.setCellValue(fs);
				}
			
				if(1== fileNum){
					wb.getSheetAt(1).removeRow(wb.getSheetAt(1).getRow(7));
				}
				String writePath = Constants.getWRITEPATH();
				new File(writePath + fileName).mkdir();
				workbook.removeSheetAt(workbook.getSheetIndex("信息采集表"));
				WriteXL.SaveAsWorkBook(workbook, Constants.getWRITEPATH()
						+ fileName + File.separator + "汇总电子版.xlsx");
//				genrPrint(workbook,titleName, Constants.getWRITEPATH() + fileName+ File.separator + "汇总打印版.xlsx");
			} else {

			}
			// 删除上传的文件
			// 返回下载文件名 须处理特殊符号 ' " /
			// System.out.println(fileName.replaceAll("/", "%"));
			wb.close();
		}

		workbook.close();
		long stop = System.currentTimeMillis();
		logger.debug("copySheet: " + (stop - start));
		if (flag) {
			logger.debug("{\"result\":\"" + fileName + "\"}");
			response.getWriter().write("{\"result\":\"" + fileName + "\"}");
		} else {
			logger.debug("{\"err\":\"错误\",\"result\":\" 请修改以下错误后重新汇总:\\n<br/>" + msg + "\"}");
			response.getWriter().write("{\"err\":\"错误\",\"result\":\" 请修改以下错误后重新汇总:\\n<br/>" + msg + "\"}");
		}
	}
	
	/**
	 * 二次合并
	 * @param request
	 * @param response
	 * @return
	 * @throws Exception
	 */
	@ResponseBody 
	@RequestMapping("/start2.do")
	public void start2(HttpServletRequest request,HttpServletResponse response) throws Exception{
		response.setContentType("application/json");
		 DecimalFormat decimalFormat=new DecimalFormat("#");
		logger.debug("进入 start2.do");
		Workbook workbook =  ReadXL.getWorkBook(Constants.getTempletPath());
		long start = System.currentTimeMillis();
		//上传的文件所在路径
		String localForld = Constants.getRealServerPath()+"upload"+File.separator+request.getParameter("localFolad");
		//所有上传的
		List<Workbook> uploadwbs = ReadXL.getWorkBookList(localForld);
		if(null == uploadwbs){
			logger.debug("{\"err\":\"错误\",\"result\":\" 请检查您的zip压缩文件！\"}");
			response.getWriter().write("{\"err\":\"错误\",\"result\":\" 请检查您的zip压缩文件！\"}");
			return;
		}
		//表头名称
		String titleName = "";
		//测试环境 忽略工作站名
		String testok =  request.getParameter("testok");
		String testStyle = request.getParameter("testStyle");
		//开始写入的位置
		int rowPos = 5;
		//文件数量
		int fileNum = uploadwbs.size();
		//sheet 页数量
		int sheetNum  = 0;
		//数据校验 错误提示信息
		String msg = "";

		// 若有校验错误 则不再继续汇总 且不生成汇总文件及打印 true为校验正确
		boolean flag = true;
		
//		int [] t =  new int[sheetNum];//单独处理的数量
		
		List<Double> t = new ArrayList<Double>();
		//工作站名称
		List<String> stationNames = new ArrayList<String>();
		//所有工作站名称
//		StringBuilder sb = new StringBuilder();
		//单独处理党委
		List<String> alldws = new ArrayList<String>();
		// 单独处理党员
//		List<String> alldys = new ArrayList<String>();
		//返回前台的文件名 
		String fileName = utils.getRandomString()+"";
		
		if (workbook instanceof HSSFWorkbook) {
			logger.debug("模版格式错误!");
		} else if (workbook instanceof XSSFWorkbook) {
			// 数字 汉字 汇总表
			Map<String, String> totalMapRec = Constants.totalMap;
			// 模版
			XSSFWorkbook wb = (XSSFWorkbook) workbook;
			
			//第一格 样式 复制时用
			XSSFCellStyle fheadStyle = wb.getSheetAt(2).getRow(0).getCell(0).getCellStyle();
			XSSFCellStyle toCellStyle = wb.createCellStyle();
			toCellStyle.setBorderBottom(CellStyle.BORDER_THIN);//下边框       
			toCellStyle.setBorderLeft(CellStyle.BORDER_THIN);//左边框       
			toCellStyle.setBorderRight(CellStyle.BORDER_THIN);//右边框       
			toCellStyle.setBorderTop(CellStyle.BORDER_THIN);//上边框    
			XSSFCellStyle headStyle =wb.createCellStyle();
			headStyle.setFont(fheadStyle.getFont());
			
			
			//变量定义
			Workbook uploadwb;
			XSSFWorkbook fromwb;
			int thissheetNum;
			Map<String, String> readBasic;
			String tosheetName;


			XSSFSheet datasheet ;
			XSSFRow row6;
			Set<String> keys;
			String value; 
			int index;
			XSSFCell cell;
			int sss =1;//复制sheet页计数
			int count = 0;
			long curTime = System.currentTimeMillis();
			//行 列 样式
			XSSFRow row5  = wb.getSheetAt(1).getRow(5);
			XSSFCellStyle hrowStyle = row5.getRowStyle();//实际为行格式  .getCell(1).getCellStyle()
//			XSSFCellStyle rowStyle = row5.getCell(1).getCellStyle();//实际的列格式
			
			Map<Integer,XSSFCellStyle> cellStylesMap = new HashMap<Integer,XSSFCellStyle>();
			int cellIndex =0;
			for(Cell c :row5){
				XSSFCellStyle cs = (XSSFCellStyle) c.getCellStyle();
				cs.setWrapText(true);
				cs.setAlignment(CellStyle.ALIGN_CENTER);
				cellStylesMap.put(cellIndex++, cs);
			}
			
			//累计行 列 样式
			XSSFRow row7  = wb.getSheetAt(1).getRow(7);
			XSSFCellStyle row7Style = row7.getRowStyle();//实际为行格式  .getCell(1).getCellStyle()
//			XSSFCellStyle rowStyle = row5.getCell(1).getCellStyle();//实际的列格式
			
			Map<Integer,XSSFCellStyle> r7cellStylesMap = new HashMap<Integer,XSSFCellStyle>();
			int r7cellIndex =0;
			for(Cell c :row7){
				XSSFCellStyle cs = (XSSFCellStyle) c.getCellStyle();
//				cs.setWrapText(true);
				r7cellStylesMap.put(r7cellIndex++, cs);
			}
			wb.getSheetAt(1).shiftRows(7, 7, -1,true,true);
			for (int i = 0; i < fileNum; i++) {
				 uploadwb = uploadwbs.get(i);
				// 应为获得模版汇总页
				if (uploadwb instanceof XSSFWorkbook) {
					// 数据采集页
					 fromwb = (XSSFWorkbook) uploadwb;
					// sheet 页数量 出去汇总页和封面
					 thissheetNum = fromwb.getNumberOfSheets() - 2;
					sheetNum += thissheetNum;
					for (int sn = 0; sn < thissheetNum; sn++) {
						// 得到采集表数据
						readBasic = ReadBasic.XSSFRead(
								fromwb, 2 + sn);
						 tosheetName = readBasic.get("商务站名称");
						 titleName = (("".equals(titleName) || titleName.equals(readBasic.get("所在区县")))?readBasic.get("所在区县"):"北京市");
						if ("checked".equals(testok))
							tosheetName += utils.geneRandomString(10);
//						if (sb.lastIndexOf(tosheetName + "reco") != -1 ) {
//							int x = sb.lastIndexOf(tosheetName + "reco");
//							if(x == 0 || "reco".equals(sb.substring(x-4))){
						if(stationNames.contains(tosheetName)){
								msg = "第"+(i+1)+"个文件中工作站 " + tosheetName + "已经存在! 请修改后重新上传!";
								response.getWriter().write("{\"err\":\"错误\",\"result\":\"" + msg + "\"}");
								uploadwb.close();
								wb.close();
								workbook.close();
								logger.debug("{\"err\":\"错误\",\"result\":\"" + msg + "\"}");
								return;
//							}
						}
//						sb = sb.append(tosheetName + "reco");
						tosheetName = tosheetName.replace("；", "");
						stationNames.add(tosheetName);

						// 进行数据校验

						// 5.AD列填报“党员数”小于等于AC列“专职工作人员”填报数
						Double dys = Double.parseDouble((null == readBasic
								.get("专职党员")
								|| "".equals(readBasic.get("专职党员")) ? "0"
								: readBasic.get("专职党员")));
						Double zzgzry = Double.parseDouble((null == readBasic
								.get("专职人员")
								|| "".equals(readBasic.get("专职人员")) ? "0"
								: readBasic.get("专职人员")));
						if (dys > zzgzry) {
							msg += "工作站： " + tosheetName + " 中 专职工作人员数("
									+ zzgzry + ") 小于其中党员数(" + dys + ")\\n<br/>";
							flag = false;
						}

						// 6.AG列填报“党建工作指导员”数小于等于AF列填报“兼职工作人员”数
						Double djzdy = Double.parseDouble((null == readBasic
								.get("党建指导员")
								|| "".equals(readBasic.get("党建指导员")) ? "0"
								: readBasic.get("党建指导员")));
						Double jzry = Double.parseDouble((null == readBasic
								.get("兼职人员")
								|| "".equals(readBasic.get("兼职人员")) ? "0"
								: readBasic.get("兼职人员")));
						if (djzdy > jzry) {
							msg += "工作站： " + tosheetName + " 中 党建工作指导员数("
									+ djzdy + ") 大于兼职工作人员数(" + jzry
									+ ")\\n<br/>";
							flag = false;
						}

						// 7.AK列“上一年度工作经费”=AL列“市级拨付”+AM列“区县拨付”+AN列“街道（乡镇）拨付”+AO列“其他途径”
						BigDecimal sndgzjf =new BigDecimal((null == readBasic
								.get("2014年经费")
								|| "".equals(readBasic.get("2014年经费")) ? "0"
								: readBasic.get("2014年经费")));
						BigDecimal sjbf = new BigDecimal
								((null == readBasic.get("市级拨付")
										|| "".equals(readBasic.get("市级拨付")) ? "0"
										: readBasic.get("市级拨付")));
						BigDecimal qxbf = new BigDecimal
								((null == readBasic.get("区县拨付")
										|| "".equals(readBasic.get("区县拨付")) ? "0"
										: readBasic.get("区县拨付")));
						BigDecimal  jdbf = new BigDecimal 
								((null == readBasic.get("街道拨付")
								|| "".equals(readBasic.get("街道拨付")) ? "0"
								: readBasic.get("街道拨付")));
						BigDecimal  qttj = new BigDecimal 
								((null == readBasic.get("其他途径")
										|| "".equals(readBasic.get("其他途径")) ? "0"
										: readBasic.get("其他途径")));

						if (!(sndgzjf.doubleValue() == (sjbf.add(qxbf).add( jdbf).add (qttj)).doubleValue())) {
							msg += "工作站： " + tosheetName + " 中 2014年工作经费(" + decimalFormat.format(sndgzjf)
									+ ") 与总来源数(" + decimalFormat.format(sjbf.add(qxbf).add( jdbf).add (qttj))
									+ ")不相等\\n<br/>";
							flag = false;
						}
						if (flag) {
							boolean copyStyleFlag  = Constants.SWITCHSTYLECOPY;
							if ("checked".equals(testStyle)) copyStyleFlag = true;
							utils.copySheetTest(wb, 2, fromwb, 2 + sn,tosheetName, toCellStyle, headStyle,copyStyleFlag);
							if(sss > 1 &&  Constants.SWITCHOVERTIME && ((System.currentTimeMillis()- curTime) > Constants.OVERTIME)){//文件处理超时 
								if(count > Constants.OVERTIMECOUNT){
									msg = "复制 "+ tosheetName+ "超时 。可能原因：1 、文件过多或过大；2、服务器压力。\\n<br/>"+ msg;
									flag = false;
									uploadwb.close();
									fromwb.close();
									break;
								}
								logger.debug("超时计数 ： "+ count++ +"第"+ sss + "个Sheet页    "+tosheetName + "复制超时");
							}
							curTime = System.currentTimeMillis();
							logger.debug("第"+i+"个文件，共"+"第"+ sss++ + "个Sheet页    "+tosheetName + "复制完成");
							// 模版汇总页
							 datasheet = wb.getSheetAt(1);
							 row6 = null;
							// if(rowPos>6){
							row6 = datasheet.createRow(rowPos);
							row6.setRowStyle(hrowStyle);
							// }else{
							// row6 = datasheet.getRow(rowPos);
							// }
							keys = totalMapRec.keySet();
							 value = "";// 采集表中的值
							for (String k : keys) {
								value = readBasic.get(k);
								 index = Integer.parseInt(totalMapRec.get(k));
								 cell = row6.getCell(index);
								 value = (null == value || "".equals(value))?"0":value;
								if (null == cell) {
									cell = row6.createCell(index);
									cell.setCellStyle(cellStylesMap.get(index));
								}
								if (k.startsWith("覆盖楼栋")) {
									t.add(Double.parseDouble(value));
									continue;
								} else if ("其他".equals(k)) {
									value = readBasic.get("其他单位");
//									cell.setCellValue(value);
									utils.setCellVall(cell,value);

								} else if ("党委".equals(k)) {
									value = readBasic.get("党委");
									value = (null == value || "".equals(value))?"0":value;
									alldws.add(value);
								} else if ("党总支".equals(k)) {
									value = readBasic.get("总支部");
//									cell.setCellValue(value);
									utils.setCellVall(cell,value);
								} else if ("党支部".equals(k)) {
									value = readBasic.get("支部");
//									cell.setCellValue(value);
									utils.setCellVall(cell,value);
								} else if ("党费返还及党组织和活动经费".equals(k)) {
									value = readBasic.get("党费返还以及党组织和活动经费");
//									cell.setCellValue(value);
									utils.setCellVall(cell,value);
								} else if ("街道经费".equals(k)) {
									value = readBasic.get("街道拨付");
//									cell.setCellValue(value);
									utils.setCellVall(cell,value);
								} else if ("".equals(k)) {

								} else {
									utils.setCellVall(cell,value);
								}
							}
							rowPos++;
						}
					}
					uploadwb.close();
					fromwb.close();
				} else {
					msg = "第"+(i+1)+"个文件格式错误，实际为xls文件!";
					response.getWriter().write("{\"err\":\"错误\",\"result\":\"" + msg + "\"}");
					logger.debug("{\"err\":\"错误\",\"result\":\"" + msg + "\"}");
					return;
				}
				uploadwb.close();
			}
			if(flag){
				wb.getSheetAt(0).getRow(4).getCell(0).setCellValue("2014年"+titleName+"商务楼宇工作站情况汇总表");
				wb.getSheetAt(1).getRow(0).getCell(0).setCellValue("2014年"+titleName+"商务楼宇工作站情况汇总表");
				
				// 工作站名连接到sheet页
				CreationHelper createHelper = workbook.getCreationHelper();  
				
				CellStyle hlink_style = wb.createCellStyle(); 
				hlink_style.setBorderBottom(CellStyle.BORDER_THIN);//下边框       
				hlink_style.setBorderLeft(CellStyle.BORDER_THIN);//左边框       
				hlink_style.setBorderRight(CellStyle.BORDER_THIN);//右边框       
				hlink_style.setBorderTop(CellStyle.BORDER_THIN);//上边框    
				hlink_style.setAlignment(CellStyle.ALIGN_CENTER);
				hlink_style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
				hlink_style.setWrapText(true);
				Font hlink_font = wb.createFont();
//				hlink_font.setUnderline(Font.U_SINGLE);  
				hlink_font.setColor(IndexedColors.BLUE.getIndex());  
				hlink_style.setFont(hlink_font);  
				
				
				//特殊处理 ( 覆盖楼宇楼栋数、 超链接)
				for(int k =0;k<sheetNum;k++){
					
					try {
						Hyperlink link = createHelper.createHyperlink(org.apache.poi.common.usermodel.Hyperlink.LINK_DOCUMENT);  
						Row spcRow = wb.getSheetAt(1).getRow(k+5);
						String sname = stationNames.get(k);//工作站名称
						spcRow.createCell(0).setCellValue(k+1);//序号
						spcRow.getCell(0).setCellStyle(cellStylesMap.get(0));
						spcRow.getCell(4).setCellValue(t.get(k));//覆盖楼宇
						link.setAddress(""+sname+"!A1");  
						spcRow.getCell(3).setHyperlink(link);  
						spcRow.getCell(3).setCellStyle(hlink_style);
						spcRow.getCell(16).setCellType(Cell.CELL_TYPE_NUMERIC);
						spcRow.getCell(16).setCellValue(Integer.parseInt(alldws.get(k)));
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
						wb.close();
						workbook.close();
						response.getWriter().write("{\"err\":\"错误\",\"result\":\"部分数据汇总失败，请检查后重新汇总！\"}");
						logger.debug("{\"err\":\"错误\",\"result\":\"部分数据汇总失败，请检查后重新汇总！\"}");
						return;
					}  
				}
				
				/*for(int i =0;i<55;i++){
					workbook.getSheetAt(1).autoSizeColumn(i,true);
				}*/
				Row row8 = wb.getSheetAt(1).createRow(5+sheetNum);//.createCell(4).setCellFormula("SUM(E6,E7)");
				row8.setRowStyle(row7Style);
				Map<Integer, String> m = Constants.forMap;
				Set<Integer> ks = m.keySet();
				CellRangeAddress cellRangeAddress = new CellRangeAddress(5+sheetNum, 5+sheetNum, 0, 3);  
			    wb.getSheetAt(1).addMergedRegion(cellRangeAddress);  
			    row8.createCell(0).setCellValue("累计");
				for(Integer k : ks){
					if(k == 0|| k == 1 || k == 2 ||k == 3){
						XSSFCellStyle tem = r7cellStylesMap.get(k);
						tem.setAlignment(CellStyle.ALIGN_CENTER);
					};
					Cell c8 =row8.createCell(k);
					c8.setCellStyle(r7cellStylesMap.get(k));
					String fs = m.get(k);
					if(fs.contains(":")){
						fs = fs.replace("7", (5+sheetNum)+"");
						fs = fs.replace("AC8=0", "AC"+(6+sheetNum)+"=0");
						fs = fs.replace("AF8=0", "AF"+(6+sheetNum)+"=0");
						fs = fs.replace("AH8=0", "AH"+(6+sheetNum)+"=0");
						fs = fs.replace("AC8,0", "AC"+(6+sheetNum)+",0");
						fs = fs.replace("AF8,0", "AF"+(6+sheetNum)+",0");
						fs = fs.replace("AH8,0", "AH"+(6+sheetNum)+",0");
						c8.setCellType(Cell.CELL_TYPE_FORMULA);
						c8.setCellFormula(fs);
					}
					c8.setCellValue(fs);
				}
				if(1== sheetNum){
					wb.getSheetAt(1).removeRow(wb.getSheetAt(1).createRow(7));
				}
				
				String writePath = Constants.getWRITEPATH();
				new File(writePath+fileName).mkdir();
				workbook.removeSheetAt(workbook.getSheetIndex("信息采集表"));
				logger.debug("用时"+(System.currentTimeMillis() - start) +" 汇总完成，开始写入,操作目录:"+fileName);
				
				
				WriteXL.SaveAsWorkBook(wb, Constants.getWRITEPATH()+fileName+File.separator+"汇总电子版.xlsx");
				
				//因打印版为汇总版，取消打印
//				genrPrint(wb,titleName, Constants.getWRITEPATH()+fileName+File.separator+"汇总打印版.xlsx");
				
				/*WriteXL writexl = new WriteXL(wb, Constants.getWRITEPATH()+fileName+File.separator+"汇总电子版.xlsx");
				writexl.run();
				if(writexl.getFinished() && "已完成".equals(writexl.getMessage())){
					logger.debug("用时"+(System.currentTimeMillis() - start) +" 写入完成，开始生成打印版,操作目录:"+fileName);
					genrPrint(wb,  Constants.getWRITEPATH()+fileName+File.separator+"汇总打印版.xlsx");
				}else if("写入超时".equals(writexl.getMessage())){
					flag = false;
					msg = "写入超时 。可能原因：1 、文件过多或过大；2、服务器压力。\\n<br/>";
					logger.debug("{\"err\":\"错误\",\"result\":\" " + msg + "\"}");
					response.getWriter().write("{\"err\":\"错误\",\"result\":\" 请修改以下错误后重新汇总:\\n<br/>" + msg + "\"}");
				}*/
					
				
			}
			//删除上传的文件
			//返回下载文件名 须处理特殊符号 ' " /
//			System.out.println(fileName.replaceAll("/", "%"));
			
			wb.close();
		}
		
		
		if (flag) {
			long stop = System.currentTimeMillis();
			logger.debug("本次汇总用时: "+ (stop -start));
			response.getWriter().write("{\"result\":\""+fileName+"\"}");
		} else {
			logger.debug("{\"err\":\"错误\",\"result\":\" 请修改以下错误后重新汇总:\\n<br/>" + msg + "\"}");
			response.getWriter().write("{\"err\":\"错误\",\"result\":\" 请修改以下错误后重新汇总:\\n<br/>" + msg + "\"}");
		}
		workbook.close();
		logger.debug("{\"result\":\""+fileName+"\"}");
		return;
	}
	
	/**
	 * 下载 
	 * @param fileName
	 * @param request
	 * @param response
	 * @return
	 * @throws Exception
	 */
	@RequestMapping("/download/{fileName}.do")
	 public ModelAndView download(@PathVariable("fileName")   
	    String fileName, HttpServletRequest request, HttpServletResponse response)   
	            throws Exception {   
	  
	        response.setContentType("text/html;charset=utf-8");   
	        request.setCharacterEncoding("UTF-8");  
	        java.io.BufferedInputStream bis = null;   
	        java.io.BufferedOutputStream bos = null;   
	        String downLoadPath = "";
	        String showName = "";
	        if(fileName.startsWith("templet")){
	        	
	        	/*String[] str = fileName.split("_");
	        	String town = ConInfo.TOWN.get(str[1]);
	        	String street = ConInfo.getStreet(str[1]).get(str[2]);
	        	showName = town+street+"工作站信息采集表.xlsx";
//	        	downLoadPath = Constants.getRealServerPath()+"templet"+File.separator+"工作站信息采集表.xlsx";
//	        	downLoadPath = getDownload(town,street);
*/	        }else if(fileName.startsWith("downPrint")){
	        	showName = "汇总打印版.xlsx";
	        	fileName = fileName.substring(9);
	        	downLoadPath = Constants.getWRITEPATH()+fileName.replaceAll("%", File.separator)+File.separator+"汇总打印版.xlsx";   
	        }else{
//	        	showName = "汇总电子版.xlsx";
	        	downLoadPath = Constants.getWRITEPATH()+fileName.replaceAll("%", File.separator)+File.separator+"汇总电子版.xlsx";
	        	String str = getDownloadFor(downLoadPath);
	        	showName = str+".xlsx";
	        }
	        logger.debug("下载路径为:"+downLoadPath);   
	        try {   
	            long fileLength = new File(downLoadPath).length();   
	            response.setContentType("application/x-msdownload;");   
//	            response.setContentType("application/vnd.ms-excel");   
	            response.setHeader("Content-disposition", "attachment; filename="  //   inline
	                    + new String(showName.getBytes("gb2312"), "ISO8859-1"));   
	            response.setHeader("Content-Length", String.valueOf(fileLength));   
	            bis = new BufferedInputStream(new FileInputStream(downLoadPath));   
	            bos = new BufferedOutputStream(response.getOutputStream());   
	            byte[] buff = new byte[2048];   
	            int bytesRead;   
	            while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {   
	                bos.write(buff, 0, bytesRead);   
	            }   
	        } catch (Exception e) {   
	            e.printStackTrace();   
	        } finally {   
	            if (bis != null)   
	                bis.close();   
	            if (bos != null)   
	                bos.close();   
	        }   
	        return null;   
	    }   
	
	/**
	 * 下载 
	 * @param fileName
	 * @param request
	 * @param response
	 * @return
	 * @throws Exception
	 */
	@RequestMapping("/downloadT.do")
	 public ModelAndView downloadT(HttpServletRequest request, HttpServletResponse response)   
	            throws Exception {   
	  
		String fileName = "";
		fileName = request.getParameter("name");
		String getT = request.getParameter("town");
		String getS = request.getParameter("street");
	        response.setContentType("text/html;charset=utf-8");   
	        request.setCharacterEncoding("UTF-8");  
	        java.io.BufferedInputStream bis = null;   
	        java.io.BufferedOutputStream bos = null;   
	        String downLoadPath = "";
	        String showName = "";
	        	
//	        	String[] str = fileName.split("_");
	        	String town = ConInfo.TOWN.get(getT);
	        	String street ="";
	        	street = request.getParameter("streetInfo");
	        	if(null==town ||"null".equals(town)){
	        		town="";
	        	}else{
	        		if(null == getS || "".equals(getS)){
	        			street = request.getParameter("streetInfo");
	        		}else{
	        			street = ConInfo.getStreet(getT).get(getS);
	
	        		}
	        		
	        	}
	        	showName = town+street+fileName+"工作站信息采集表.xlsx";
//	        	downLoadPath = Constants.getRealServerPath()+"templet"+File.separator+"工作站信息采集表.xlsx";
	        	downLoadPath = getDownload(town,street,fileName);
	        
	        logger.debug("下载路径为:"+downLoadPath);   
	        try {   
	            long fileLength = new File(downLoadPath).length();   
	            response.setContentType("application/x-msdownload;");   
//	            response.setContentType("application/vnd.ms-excel");   
	            response.setHeader("Content-disposition", "attachment; filename="  //   inline
	                    + new String(showName.getBytes("gb2312"), "ISO8859-1"));   
	            response.setHeader("Content-Length", String.valueOf(fileLength));   
	            bis = new BufferedInputStream(new FileInputStream(downLoadPath));   
	            bos = new BufferedOutputStream(response.getOutputStream());   
	            byte[] buff = new byte[2048];   
	            int bytesRead;   
	            while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {   
	                bos.write(buff, 0, bytesRead);   
	            }   
	        } catch (Exception e) {   
	            e.printStackTrace();   
	        } finally {   
	            if (bis != null)   
	                bis.close();   
	            if (bos != null)   
	                bos.close();   
	        }   
	        return null;   
	    }   
	
	/**
	 * 生成打印版
	 * @param workbook
	 * @param path
	 * @throws IOException 
	 */
	public void genrPrint(Workbook workbook,String titleName,String path) throws IOException{
		Workbook workbooktop =  ReadXL.getWorkBook(Constants.getPrintTempletPath());
		
		Sheet sheet = workbook.getSheetAt(1);//读取的表
		Map<String,String> p1Map = Constants.p1Map;
//		Map<String,String> p2Map = Constants.p2Map;
//		Map<String,String> p3Map = Constants.p3Map;
//		Map<String,String> p4Map = Constants.p4Map;
		List<Map<String,String>> maps = new ArrayList<Map<String,String>>();
		maps.add(p1Map);
//		maps.add(p2Map);
//		maps.add(p3Map);
//		maps.add(p4Map);
		Map<String,String> totalMap = Constants.totalMap;
		int pos = 5;//开始写入数据的位置
		int postor = 5;//开始读取数据的位置
		for(int m=0;m<maps.size();m++){
			Sheet sheettow = workbooktop.getSheetAt(m);//要写入的表
			CellStyle  cellstyle = sheettow.getRow(5).getCell(1).getCellStyle();
			cellstyle.setBorderBottom(CellStyle.BORDER_THIN);
//			Set<String> p1sets = p1Map.keySet();
			Set<String> p1sets = maps.get(m).keySet();
			for(int i =0;i<workbook.getNumberOfSheets()-2;i++){
				Row p1row = sheettow.createRow(i+pos);
				Cell c0 = p1row.createCell(0);
				c0.setCellStyle(cellstyle);
				c0.setCellValue(i+1);
				/*Cell ce = p1row.createCell(16);
				ce.setCellStyle(cellstyle);
				ce.setCellValue("");*/
				for(String key : p1sets){
					int cellpos = Integer.parseInt(maps.get(m).get(key));//写入列的位置
					Cell p1cell = p1row.createCell(cellpos);
					if("负责人政治面貌".equals(key)){
						key = "站负责人政治面貌";
					}else if("负责人姓名".equals(key)){
						key = "站负责人姓名";
					}else if("若是，党组织设置形式".equals(key)){
						key = "党组织设置形式";
					}else if("党员".equals(key)){
						key = (m==1)?"驻厦单位党员情况党员人数":"专职党员";
					}
					
					if(null == totalMap.get(key)) continue;
					p1cell.setCellStyle(cellstyle);
					int cellpostor = Integer.parseInt(totalMap.get(key));//读取列的位置
					
					
					Cell cell = null;
					Row readrow = sheet.getRow(postor+i);
					if(null == readrow) continue;
					cell = sheet.getRow(postor+i).getCell(cellpostor);
					if(null == cell) continue;
					int t = cell.getCellType();
					String value = "";
					DecimalFormat df = new DecimalFormat("0.0");//保1位小数且不用科学计数法
					switch (t) {
					case 0:
						/*value = cell.getNumericCellValue() + "";
						value = value.endsWith(".00")?value.substring(0,value.lastIndexOf(".")):value;*/
						value = df.format(new Double(cell.getNumericCellValue()))+"";
						value = value.endsWith(".0")?value.substring(0,value.lastIndexOf(".")):value;
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
						value = cell.getErrorCellValue()+"";
						break;
					}
					p1cell.setCellValue(value);
				}
			}
			
			for(int i =0;i<55;i++){
				sheettow.autoSizeColumn(i,true);
			}
			sheettow.getRow(0).getCell(0).setCellValue("2014年"+titleName+"商务楼宇工作站情况汇总表");
		}
		workbooktop.setFirstVisibleTab(0);
		WriteXL.SaveAsWorkBook(workbooktop, path);
		workbooktop.close();
	}

	/**
	 * 判断文件是否存在不存在则生成
	 * @param str
	 * 参数为 区县+街道
	 * @throws IOException 
	 */
	public String getDownload(String town,String street,String fileName) throws IOException{
		String filePath = Constants.getGeneDownload()+town+street+"工作站信息采集表.xlsx";
		
		
			XSSFWorkbook wb = (XSSFWorkbook) Constants.getWorkBook(Constants.getTempletDate());
			/*logger.debug("获得的workbook对象是：" + wb);
			logger.debug("获得的sheet对象是：" + wb.getSheetAt(0));
			logger.debug("获得的getRow对象是：" + wb.getSheetAt(0).getRow(0));
			logger.debug("获得的getCell 对象是：" + wb.getSheetAt(0).getRow(0).getCell(0));*/
			wb.getSheetAt(0).getRow(0).getCell(0).setCellValue("2014年"+town+street+fileName+"商务楼宇工作站情况采集表");
			wb.getSheetAt(0).getRow(4).getCell(1).setCellValue(town);
			wb.getSheetAt(0).getRow(4).getCell(3).setCellValue(street);
			wb.getSheetAt(0).getRow(3).getCell(1).setCellValue(fileName);
			WriteXL.SaveAsWorkBook(wb, filePath);
			wb.close();
		
		return filePath;
	}
	/**
	 * 对下载的文件更改显示名
	 * @param str
	 * 参数为 区县+街道
	 * @throws IOException 
	 */
	public String getDownloadFor(String filePath) throws IOException{
//		String filePath = Constants.getGeneDownload()+town+street+"工作站信息采集表.xlsx";
		
		
			XSSFWorkbook wb = (XSSFWorkbook) Constants.getWorkBook(filePath);	
			String str = wb.getSheetAt(1).getRow(0).getCell(0).getStringCellValue();
//			String town = wb.getSheetAt(0).getRow(4).getCell(1).getStringCellValue();
//			String street = wb.getSheetAt(0).getRow(4).getCell(3).getStringCellValue();
			wb.close();
		
		return str;
	}
}

