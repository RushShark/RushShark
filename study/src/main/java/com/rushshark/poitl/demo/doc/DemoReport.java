package com.rushshark.poitl.demo.doc;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.ConfigureBuilder;
import com.rushshark.poitl.demo.utils.PoiTlUtil;
import org.springframework.stereotype.Component;
import org.springframework.util.StringUtils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * 报告示例
 * User:  huyunlong
 * Date:  2023/7/13
 */
@Component
public class DemoReport {
	/**
	 * 获取一个基础配置
	 */
	private ConfigureBuilder builder;
	/**
	 * 数据模型
	 */
	private Map<String, Object> map;

	public DemoReport() {
		this.builder = PoiTlUtil.getDemoBuilder();
	}

	public DemoReport(ConfigureBuilder builder, HashMap<String, Object> map) {
		this.builder = builder;
		this.map = map;
	}

	/**
	 * 构建文档
	 * @param templateFile 模板文件路径名
	 * @param outputDoc 输出文档路径名
	 * @param builder 配置构造器
	 * @param map 数据模型
	 */
	public void buildDemoDoc(String templateFile,
							 String outputDoc,
							 ConfigureBuilder builder,
							 Map<String, Object> map) {
		//参数校验
		if (!StringUtils.hasLength(templateFile) || !StringUtils.hasLength(outputDoc)){
			//todo
			return;
		}
		if (null == map){
			//todo
			return;
		}

		//编译
		XWPFTemplate template = XWPFTemplate
				.compile(templateFile, builder.build())
				.render(map);


		//输出
		try {
			template.writeAndClose(new FileOutputStream(outputDoc));
		} catch (IOException e) {
			//todo
			e.printStackTrace();
		}
	}

	/**
	 * 构建文档
	 * @param templateFile 模板文件路径名
	 * @param outputDoc 输出文档路径名
	 * @param map 数据模型
	 */
	public void buildDemoDoc(String templateFile, String outputDoc, Map<String, Object> map){
		buildDemoDoc(templateFile, outputDoc, this.builder, map);
	}
}
