package com.jacob.office;

import java.util.Observable;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * office 转换类
 * 
 * @author alex.yao
 *
 */
public abstract class Converter extends Observable implements Runnable {

	public static final Logger log = LoggerFactory.getLogger(Converter.class);
	
	public static final int DOC_FORMAT_PDF = 17;  
	public static final int XLS_FORMAT_PDF = 0;  
	public static final int PPT_FORMAT_PDF = 32;
    
	private String inputFile; // 传入文件
	private String outputFile; // 转出文件
	
	private volatile int state;
    private static final int NEW          = 0;
    private static final int DONE         = 1;

	/**
	 * 转换方法
	 * 
	 * @param inputFile
	 * @param outputFile
	 * @param timeout
	 * @return
	 */
	public abstract boolean convertToPdf(String inputFile, String outputFile) throws Exception;
	
	public boolean isDone() {
		return state != NEW;
	}

	@Override
	public void run() {
		log.info("Begin convert task");
		try {
			boolean ret = convertToPdf(this.inputFile, this.outputFile);
			if (!ret) {
				log.error("convert pdf unsuccess...");
			} else {
				setChanged();
				notifyObservers();
			}
			state = DONE;
		} catch (Exception e) {
			e.printStackTrace();
			log.error("convert office to pdf error::" + e.getMessage());
		}
		log.info("Convert office to pdf finish.");
	}

	public String getInputFile() {
		return inputFile;
	}

	public void setInputFile(String inputFile) {
		this.inputFile = inputFile;
	}

	public String getOutputFile() {
		return outputFile;
	}

	public void setOutputFile(String outputFile) {
		this.outputFile = outputFile;
	}

}
