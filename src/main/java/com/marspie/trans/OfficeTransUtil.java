package com.marspie.trans;

import java.io.File;
import java.util.LinkedList;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.jacob.office.Converter;
import com.jacob.office.ExcelConverter;
import com.jacob.office.PPTConverter;
import com.jacob.office.WordConverter;
import com.marspie.rmi.office.OfficeType;

/**
 * office 转换工具类
 * 
 * @author alex.yao
 *
 */
public class OfficeTransUtil {
	private static final Logger log = LoggerFactory.getLogger(OfficeTransUtil.class);
	private static OfficeTransUtil instance;
	private long timeOut = 120000L;
	private ExecutorService executor;	//线程池
	private OfficeTransThread officeTransThread;
	private final LinkedList<Object[]> fileQueue = new LinkedList<Object[]>();

	private Object lock = new Object();

	public static OfficeTransUtil getInstance() {
		if (instance == null) {
			instance = new OfficeTransUtil();
			instance.init();
		}
		return instance;
	}

	private synchronized void init() {
		if (this.executor == null) {
			this.executor = Executors.newFixedThreadPool(1);
		}
		if (this.officeTransThread == null) {
			this.officeTransThread = new OfficeTransThread();
			this.officeTransThread.start();
		}
	}
	
	/**
	 * 获取转换类
	 * @param relName
	 * @param inputFile
	 * @param outputFile
	 * @param param
	 * @return
	 */
	private Converter getConverter(String inputFile, String outputFile, String officeType) {
		if("EXCEL".equals(officeType))
			return new ExcelConverter(inputFile, outputFile);
		else if("WORD".equals(officeType)) {
			return new WordConverter(inputFile, outputFile);
		} else if("PPT".equals(officeType)) {
			return new PPTConverter(inputFile, outputFile);
		}
		return null;
	}
	
	private boolean convertToPdf(String inputFile, String outputFile, String officeType) {
		long t1 = System.currentTimeMillis();
		Converter converter = getConverter(inputFile, outputFile, officeType);
		if (converter == null) {
			log.error("InputFile type is not support, inputFile::{} officeType::{}", inputFile, officeType);
			return false;
		}
		if (this.executor == null) {
			this.executor = Executors.newFixedThreadPool(1);
		}
		this.executor.execute(converter);
		while (!converter.isDone()) {
			long t2 = System.currentTimeMillis();
			if (t2 - t1 >= this.timeOut) {
				System.runFinalization();
				try {
					log.error("Convert timeout...use :: {} but not convert success!", t2 - t1);
				} catch (Exception e) {
					e.printStackTrace();
				}
				return false;
			}
			
		}
		return true;
	}

	/**
	 * 执行队列添加执行对象
	 * @param officeFilePath 源文件
	 * @param outputFilepath 目标文件
	 * @param officeType 源文件类型
	 */
	public void transform(String officeFilePath, String outputFilepath, OfficeType officeType) {
		synchronized (this.lock) {
			Object[] o = {officeFilePath, outputFilepath, officeType + ""};
			this.fileQueue.add(o);
		}
	}

	protected Object[] getNext() {
		synchronized (this.lock) {
			if (!this.fileQueue.isEmpty()) {
				return (Object[]) this.fileQueue.remove();
			}
			return null;
		}
	}
	
	public void finalize() throws Throwable {
		super.finalize();
		try {
			this.executor.shutdown();
			this.officeTransThread.stopThread();
		} catch (Throwable e) {
			log.error(e.getMessage());
			e.printStackTrace();
		}
		log.info("OfficeTrans关闭");
		System.out.println("OfficeTrans关闭");
	}
	  
	/**
	 * office 转换线程
	 * @author alex.yao
	 *
	 */
	class OfficeTransThread extends Thread {
		private boolean running = true;

		OfficeTransThread() {
		}

		public synchronized void start() {
			super.setName("OfficeTrans");
			super.start();
		}

		public void stopThread() {
			this.running = false;
		}

		@SuppressWarnings("all")
		public void run() {
			while (this.running) {
				try {
					Object[] o = OfficeTransUtil.this.getNext();
					if (o == null)
						continue;

					String src = (String) o[0];
					String target = (String) o[1];
					String officeType = (String) o[2];
					long startTime = System.currentTimeMillis();
					OfficeTransUtil.log.info("src:" + src);
					OfficeTransUtil.log.info("target:" + target);
					File targetFile = new File(target);
					if (!targetFile.exists()) {
						boolean result = OfficeTransUtil.this.convertToPdf(src, target, officeType);
						OfficeTransUtil.log.info("Convert result::" + result);
						OfficeTransUtil.log.info("Convert finish used time::" + (System.currentTimeMillis() - startTime) + " ms");
					} else {
						OfficeTransUtil.log.info(src + "->" + target + "...end:exist!!!!");
					}
				} catch (Throwable e) {
					OfficeTransUtil.log.error("", e);
				}
				try {
					Thread.sleep(2000L);
				} catch (Throwable e) {
					e.printStackTrace();
					OfficeTransUtil.log.error(e.getMessage());
				}
			}
		}
	}
}
