package com.marspie.rmi.launcher;

import org.slf4j.Logger;  
import org.slf4j.LoggerFactory; 
import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;

/**
 * 服务启动类
 * @author alex.yao
 *
 */
public class ServerLauncher {
	
	private static Logger log = LoggerFactory.getLogger(ServerLauncher.class);  
	
	@SuppressWarnings("all")
	public static void main(String[] args) {
		ApplicationContext context = new ClassPathXmlApplicationContext("classpath*:server-config.xml");
		log.info("Marspie office transfer server start.");
	}
}
