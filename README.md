## Office转PDF RMI服务
JACOB一个Java-COM中间件.通过这个组件可以在Java应用程序中调用COM组件和Win32程序库，本应用使用JACOB转换Office文档为PDF，前提系统需安装Office或WPS
[jacob 下载地址](https://sourceforge.net/projects/jacob-project/)

### 工具
1. jdk1.7+
2. maven3.0+
3. jacob-1.18

### 项目部署
1. 将lib目录下 jacob-1.18.dll 复制至jdk安装目录jdk/jre/bin/目录下
2. 将jacob.jar安装至maven本地仓库
``` bash
mvn install:install-file -Dfile=jacob.jar -DgroupId=com.jacob -DartifactId=jacob -Dversion=1.18 -Dpackaging=jar
``` 
3. 执行项目下package.bat打包项目
4. 启动marspie-officetrans服务
``` bash
java -jar target/marspie-officetrans-jar-with-dependencies.jar
```
5. 调用转换服务
``` bash
public static void main(String[] args) throws Exception {
		OfficeTransService server = ClientUtil.getProxy("127.0.0.1", 1199, "OfficeTransService", OfficeTransService.class);
		server.officeToPdf("d:\\zyzd.xlsx", "d:\\zds2.xlsx.pdf", OfficeType.EXCEL);
		System.out.println("finish!");
	}
```

