package com.jacob.office;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class WordConverter extends Converter {

	public WordConverter(String inputFile, String outputFile) {
		setInputFile(inputFile);
		setOutputFile(outputFile);
	}
	
	@Override
	public boolean convertToPdf(String inputFile, String outputFile) throws Exception {
		ActiveXComponent app = null;
		Dispatch doc = null;
		try {
			ComThread.InitSTA();
			app = new ActiveXComponent("Word.Application");
			app.setProperty("Visible", false);
			Dispatch docs = app.getProperty("Documents").toDispatch();
			doc = Dispatch.invoke(docs, "Open", Dispatch.Method, new Object[] { inputFile, new Variant(false),
					new Variant(true), new Variant(false), new Variant("pwd") }, new int[1]).toDispatch();
			Dispatch.put(doc, "RemovePersonalInformation", false);
			Dispatch.call(doc, "ExportAsFixedFormat", outputFile, DOC_FORMAT_PDF);
		} catch (Exception e) {
			e.printStackTrace();
			log.error("Word Convert to pdf error source::{} target::{} error::{}", inputFile, outputFile, e.getMessage());
			return false;
		} finally {
			if (doc != null) {
				Dispatch.call(doc, "Close", false);
			}
			if (app != null) {
				app.invoke("Quit", 0);
			}
			ComThread.Release();
		}
		return true;
	}
}
