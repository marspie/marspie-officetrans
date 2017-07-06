package com.jacob.office;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;

public class PPTConverter extends Converter {

	public PPTConverter(String inputFile, String outputFile) {
		setInputFile(inputFile);
		setOutputFile(outputFile);
	}
	
	@Override
	public boolean convertToPdf(String inputFile, String outputFile) throws Exception {
		ActiveXComponent app = null;
		Dispatch ppts = null;
		Dispatch ppt = null;
		try {
			app = new ActiveXComponent("PowerPoint.Application");
			ppts = app.getProperty("Presentations").toDispatch();
			ppt = Dispatch.call(ppts, "Open", inputFile,
										true,//ReadOnly
										true,//Untitled指定文件是否有标题
										false//WithWindow指定文件是否可见
										).toDispatch();
			
			Dispatch.call(ppt, "SaveAs", outputFile, PPT_FORMAT_PDF);
		} catch (Exception e) {
			e.printStackTrace();
			log.error("PPT Convert to pdf error source::{} target::{} error::{}", inputFile, outputFile, e.getMessage());
			return false;
		} finally {
			if (ppt != null) {
				Dispatch.call(ppt, "Close", false);
			}
			if (app != null) {
				app.invoke("Quit", 0);
			}
			ComThread.Release();
		}
		return true;
	}
}
