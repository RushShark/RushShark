package com.rushshark.poltl.demo;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;
import com.deepoove.poi.data.HyperlinkTextRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.Texts;
import com.rushshark.poltl.demo.utils.PoiTlUtil;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

/**
 * User:  huyunlong
 * Date:  2023/7/12
 */
public class Demo1 {
	public static void main(String[] args) throws IOException {
		//获取一个基础配置
		ConfigureBuilder builder = PoiTlUtil.getDemoBuilder();

		//数据模型
		HashMap<String, Object> map = new HashMap<>();

		//*********文本*************
		//String ：文本
		//		//TextRenderData ：有样式的文本
		//		map.put("author", Texts.of("Sayi").color("000000").create());
		//		//HyperlinkTextRenderData ：超链接和锚点文本
		//		map.put("link", Texts.of("website").link("http://deepoove.com").create());
		//		//Object ：调用 toString() 方法转化为文本
		//		map.put("anchor", Texts.of("anchortxt").anchor("appendix1").create());
		map.put("orgName", "xxx医院");
		map.put("begMonth", "2023.01");
		map.put("endMonth", "2023.07");
		map.put("reportDate", "2023.07.12");



		//***列表

		//***图片

		//***表格

		//***图表


		//编译
		XWPFTemplate template = XWPFTemplate
				.compile("C:\\Users\\huyunlong\\Desktop\\DIP运营分析报告-报告模板.docx", builder.build())
				.render(map);


		//输出
		template.writeAndClose(new FileOutputStream("C:\\Users\\huyunlong\\Desktop\\郑州蓝博电子技术有限公司_DIP运营分析报告.docx"));
	}
}
