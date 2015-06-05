package com.lydj.Controller.read;
import java.io.BufferedOutputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.lydj.Controller.util.Constants;


public class WriteXL extends Thread{
	
	//是否写入完成
	boolean isFinished = false;
	//超时时间
	long  exprTime = Constants.OVERWRITETIME;
	boolean switchExprTime = Constants.SWITCHWRITE;
    Workbook workbook;
	String savePath;
	String message ="已完成";
	
	public WriteXL(String message) {
		super();
		this.message = message;
	}




	public String getMessage() {
		return message;
	}




	public void setMessage(String message) {
		this.message = message;
	}




	public WriteXL() {
		super();
	}


	

	public WriteXL(Workbook workbook, String savePath) {
		super();
		this.workbook = workbook;
		this.savePath = savePath;
	}




	public WriteXL(boolean isFinished, long exprTime, boolean switchExprTime) {
		super();
		this.isFinished = isFinished;
		this.exprTime = exprTime;
		this.switchExprTime = switchExprTime;
	}


	



	public boolean getFinished() {
		return isFinished;
	}




	public void setFinished(boolean isFinished) {
		this.isFinished = isFinished;
	}




	public long getExprTime() {
		return exprTime;
	}




	public void setExprTime(long exprTime) {
		this.exprTime = exprTime;
	}




	public boolean getSwitchExprTime() {
		return switchExprTime;
	}




	public void setSwitchExprTime(boolean switchExprTime) {
		this.switchExprTime = switchExprTime;
	}




	public Workbook getWorkbook() {
		return workbook;
	}




	public void setWorkbook(Workbook workbook) {
		this.workbook = workbook;
	}




	public String getSavePath() {
		return savePath;
	}




	public void setSavePath(String savePath) {
		this.savePath = savePath;
	}




	/** Excel文件的存放位置。注意是正斜线 */
	public static void SaveAsWorkBook(Workbook workbook,String savePath) {
		// 新建一输出文件流
		FileOutputStream fOut;
		try {
			fOut = new FileOutputStream(savePath);
			BufferedOutputStream bos = new BufferedOutputStream(fOut);
			workbook.setForceFormulaRecalculation(true);
			workbook.write(bos);
			System.gc();
//			fOut.flush();
			// 操作结束，关闭文件
			fOut.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/** Excel文件的存放位置。注意是正斜线 */
	public   void SaveAsWorkBook2() {
		// 新建一输出文件流
		FileOutputStream fOut;
		WriteListener lis = new WriteListener(this);
		SXSSFWorkbook sf = new SXSSFWorkbook((XSSFWorkbook) workbook,30);
		sf.setCompressTempFiles(true);
		try {
			fOut = new FileOutputStream(savePath);
			BufferedOutputStream bos = new BufferedOutputStream(fOut);
			if(this.getSwitchExprTime()) {
				lis.start();
			}
			System.out.println("开始写入!");
//			workbook.write(bos);
			sf.write(bos);
//			fOut.flush();
			// 操作结束，关闭文件
			fOut.close();
			sf.dispose();
			this.setFinished(true);
			lis.interrupt();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}




	@Override
	public void run() {
		// TODO Auto-generated method stub
		synchronized (workbook){
			SaveAsWorkBook2();
		}
	}

	
	
	
}