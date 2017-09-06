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
        Dispatch ppt = null;
        try {
            ComThread.InitSTA();
            app = new ActiveXComponent("PowerPoint.Application");
            Dispatch ppts = app.getProperty("Presentations").toDispatch();
            /*
             * call 
             * param 4: ReadOnly
             * param 5: Untitled指定文件是否有标题
             * param 6: WithWindow指定文件是否可见
             * */
            ppt = Dispatch.call(ppts, "Open", inputFile, true,true, false).toDispatch();
            Dispatch.call(ppt, "SaveAs", outputFile, PPT_FORMAT_PDF); // ppSaveAsPDF为特定值32
        } catch (Exception e) {
            e.printStackTrace();
            log.error("PPT Convert to pdf error source::{} target::{} error::{}", inputFile, outputFile, e.getMessage());
			return false;
        } finally {
            if (ppt != null) {
                Dispatch.call(ppt, "Close");
            }
            if (app != null) {
                app.invoke("Quit");
            }
            ComThread.Release();
        }
		return true;
	}
}
