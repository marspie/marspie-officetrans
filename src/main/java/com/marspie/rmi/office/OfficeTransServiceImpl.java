package com.marspie.rmi.office;

import java.io.File;

import org.springframework.stereotype.Service;

import com.marspie.trans.OfficeTransUtil;

@Service("officeTransServiceImpl")
public class OfficeTransServiceImpl implements OfficeTransService {

	@Override
	public void officeToPdf(String sourcePath, String targetPath, OfficeType officeType) {
		if(sourcePath == null || !new File(sourcePath).exists()){
			log.error("Source file is not exists. source::{}", sourcePath);
			return;
		}
		log.info("Add Convert task source::{} target::{} officeType::{}", sourcePath, targetPath, officeType);
		OfficeTransUtil.getInstance().transform(sourcePath, targetPath, officeType);   
	}
}
