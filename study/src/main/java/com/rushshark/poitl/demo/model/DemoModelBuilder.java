package com.rushshark.poitl.demo.model;

import com.deepoove.poi.data.Texts;
import org.springframework.stereotype.Component;

import java.util.HashMap;
import java.util.Map;

/**
 * 构建示例数据模型
 * User:  huyunlong
 * Date:  2023/7/13
 */
@Component
public class DemoModelBuilder {

	/**
	 * 构建数据模型
	 * @return
	 */
	public Map<String, Object> buildData(){
		Map<String, Object> map = new HashMap<>();
		//*********文本*************
		//String ：文本
		//		//TextRenderData ：有样式的文本
		//		map.put("author", Texts.of("Sayi").color("000000").create());
		//		//HyperlinkTextRenderData ：超链接和锚点文本
		//		map.put("link", Texts.of("website").link("http://deepoove.com").create());
		//		//Object ：调用 toString() 方法转化为文本
		//		map.put("anchor", Texts.of("anchortxt").anchor("appendix1").create());
		this.putText(map);




		//***列表

		//***图片

		//***表格

		//***图表
		return map;
	}
	private void putText(Map<String, Object> map){
		map.put("orgName", "xxx医院");
		map.put("begMonth", "2023.01");
		map.put("endMonth", "2023.07");
		map.put("setlCaseNum", "1555");
		map.put("totalFee", "1000");
		map.put("dipFee", "1200");
		map.put("feePayrate", "120");
		map.put("medInsFun", "88");
		map.put("dipFun", "666");
		map.put("funDiff", "5");
		map.put("funRate", "12");
		map.put("feeDiff", "92000");
		map.put("avgFee", "92000");
		map.put("dipAvgFee", "92000");
		map.put("avgFeeDiff", Texts.of("92000").bookmark("ss").create());
		map.put("avgLos", "92000");

		map.put("setlCaseNum", Texts.of("1555")
				.color("FDC972").bold().fontSize(20)
				.anchor("ss").create());
	}
}
