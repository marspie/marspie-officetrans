package com.marspie.rmi.client.util;

import org.springframework.remoting.rmi.RmiProxyFactoryBean;

import com.marspie.rmi.office.OfficeTransService;
import com.marspie.rmi.office.OfficeType;

public class ClientUtil
{
  public static <T> T getProxy(String host, int port, String serviceName, Class<T> serviceInterface)
    throws Exception
  {
    String url = "rmi://" + host + ":" + port + "/" + serviceName;
    return getProxy(url, serviceInterface);
  }

  public static <T> T getProxy(String serviceUrl, Class<T> serviceInterface)
    throws Exception
  {
    RmiProxyFactoryBean factory = new RmiProxyFactoryBean();
    factory.setServiceUrl(serviceUrl);
    factory.setServiceInterface(serviceInterface);
    factory.afterPropertiesSet();

    return serviceInterface.cast(factory.getObject());
  }

	public static void main(String[] args) throws Exception {
		OfficeTransService s = ClientUtil.getProxy("127.0.0.1", 1199, "OfficeTransService", OfficeTransService.class);
		s.officeToPdf("d:\\zyzd.xlsx", "d:\\zds2.xlsx.pdf", OfficeType.EXCEL);
		System.out.println("finish!");
	}
}