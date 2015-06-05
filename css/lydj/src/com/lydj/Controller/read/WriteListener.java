package com.lydj.Controller.read;

public class WriteListener extends Thread {

	
	WriteXL writexl;
	
	
	public WriteListener() {
		super();
	}


	public WriteListener(WriteXL writexl) {
		super();
		this.writexl = writexl;
	}


	@Override
	public void run() {
		// TODO Auto-generated method stub
		try {
			System.out.println("监听已启动"+Thread.currentThread().getName() +writexl.getExprTime());
			Thread.sleep(writexl.getExprTime());
			if(!writexl.getFinished()){
				System.out.println("写入超时");
				writexl.setMessage("写入超时");
				writexl.interrupt();
			}
			
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
//			e.printStackTrace();
			System.out.println("写入完成，监听取消");
		}
		
	}

}
