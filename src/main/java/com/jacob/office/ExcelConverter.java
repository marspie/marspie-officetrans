package com.jacob.office;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;

/**
 * Excel文件转换
 * @author alex.yao
 *
 */
public class ExcelConverter extends Converter {
	
	public ExcelConverter(String inputFile, String outputFile) {
		setInputFile(inputFile);
		setOutputFile(outputFile);
	}

	@Override
	public boolean convertToPdf(String inputFile, String outputFile) throws Exception {
		ActiveXComponent app = null;
		Dispatch excel = null;
		try {
			ComThread.InitSTA();
			app = new ActiveXComponent("Excel.Application");
			app.setProperty("Visible", false);
			Dispatch excels = app.getProperty("Workbooks").toDispatch();
			excel = Dispatch.call(excels, "Open", inputFile, false, true).toDispatch();
			Dispatch.call(excel, "ExportAsFixedFormat", XLS_FORMAT_PDF, outputFile);
		} catch (Exception e) {
			e.printStackTrace();
			log.error("Excel Convert to pdf error source::{} target::{} error::{}", inputFile, outputFile, e.getMessage());
			return false;
		} finally {
			if (excel != null) {
				Dispatch.call(excel, "Close", false);
			}
			if (app != null) {
				app.invoke("Quit");
			}
			ComThread.Release();
		}
		return true;
	}

}
