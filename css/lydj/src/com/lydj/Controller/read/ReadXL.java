package com.lydj.Controller.read;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.lydj.Controller.util.utils;

public class ReadXL {
	/** Excel文件的存放位置。注意是正斜线 */
//	public static String fileToBeRead = Constants.TEMPLET;
	public static Workbook getWorkBook(String path){
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
		}  catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return workbook;
		
	}
	
	/**
	 * 目录下所有文件
	 * @param path
	 * @return
	 */
	public static List<Workbook> getWorkBookList(String forld){
		List<Workbook> list = new ArrayList<Workbook>();
		File file = new File(forld);
		if(!utils.zipDeco(file)) return null;
		File [] files = file.listFiles();
		
		for(File f :files){
			if( ! f.getName().contains(".xlsx")){
				continue;
			}
			Workbook workbook = null;
			try {
//				workbook = new HSSFWorkbook(new FileInputStream(
//						f));
				FileInputStream fis = new FileInputStream(f);
				 workbook = new HSSFWorkbook(fis);
				 fis.close();
			} catch (OfficeXmlFileException e) {
				// TODO Auto-generated catch block
				try {
					FileInputStream fis = new FileInputStream(f);
					workbook = new XSSFWorkbook(fis);
					fis.close();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					return null;
				} 
			}  catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			list.add(workbook);
		}
		
		return list;
		
	}
}