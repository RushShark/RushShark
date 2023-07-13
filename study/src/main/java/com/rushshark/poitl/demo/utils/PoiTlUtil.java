package com.rushshark.poitl.demo.utils;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;
import com.deepoove.poi.util.RegexUtils;

import java.util.List;
import java.util.stream.Collectors;

/**
 * Poi-Tl 工具类，详细使用见 http://deepoove.com/poi-tl/
 * User:  huyunlong
 * Date:  2023/7/12
 */
public class PoiTlUtil {
	/**
	 * 获取{{}}中的内容表达式
	 */
	public static final String REG_TL = "\\{\\{([^}]*)\\}\\}";
	/**
	 * 	获取一个全局配置
	 * @return
	 */
	public static ConfigureBuilder getDemoBuilder(){

		//poi-tl提供了类 Configure 来配置常用的设置
		ConfigureBuilder builder = Configure.builder();

		//允许内容为除了标签前后缀外的任意字符
		builder.buildGrammerRegex(RegexUtils.createGeneral("{{", "}}"));

		//使用SpringEL表达式
		/*
		常用EL表达式，更多见 https://docs.spring.io/spring-framework/docs/5.3.18/reference/html/core.html#expressions
		{{name}}
		{{name.toUpperCase()}}
		{{name == 'poi-tl'}}
		{{empty?:'这个字段为空'}}
		{{sex ? '男' : '女'}}
		{{new java.text.SimpleDateFormat('yyyy-MM-dd HH:mm:ss').format(time)}}
		{{price/10000 + '万元'}}
		{{dogs[0].name}}
		{{localDate.format(T(java.time.format.DateTimeFormatter).ofPattern('yyyy年MM月dd日'))}}
		 */
		builder.useSpringEL();

		/*
		数据模型支持JSON字符串序列化，可以方便的构造远程HTTP或者RPC服务，需要引入相应依赖
			<dependency>
				<groupId>com.deepoove</groupId>
				<artifactId>poi-tl-jsonmodel-support</artifactId>
				<version>1.0.0</version>
			</dependency>
		 */
		//builder.addPreRenderDataCastor(new GsonPreRenderDataCastor());

		//poi-tl可以在发生这种错误时对计算结果进行配置，默认会认为标签值为null。当我们需要严格校验模板是否有人为失误时，可以抛出异常：
		//builder.useDefaultEL(true);

		//如果使用SpringEL表达式，可以通过参数来配置是否抛出异常
		builder.useSpringEL(false);

		//poi-tl默认的行为会清空标签，如果希望对标签不作任何处理：
		//builder.setValidErrorHandler(DiscardHandler());

		//如果希望执行严格的校验，直接抛出异常：
		//builder.setValidErrorHandler(AbortHandler());

		return builder;
	}

	/**
	 * 获取word文档中所有Poi_tl标签
	 * @param fileName 文件名
	 * @return
	 */
	public static List<String> getLabel(String fileName) {
		XWPFTemplate template = XWPFTemplate.compile(fileName);
		return template.getElementTemplates().stream().map(e -> e.variable()).collect(Collectors.toList());
	}
}
