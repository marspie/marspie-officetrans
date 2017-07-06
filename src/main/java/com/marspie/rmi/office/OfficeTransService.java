package com.marspie.rmi.office;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public interface OfficeTransService {
	
	public Logger log = LoggerFactory.getLogger(OfficeTransService.class);

	/**
	 * 转换PDF
	 * @param sourcePath 源文件
	 * @param targetPath 目标PDF文件
	 * @param params 参数
	 */
	public void officeToPdf(String sourcePath, String targetPath, OfficeType officeType);
	    
}
