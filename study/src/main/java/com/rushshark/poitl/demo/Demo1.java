package com.rushshark.poitl.demo;

import com.rushshark.poitl.demo.utils.PoiTlUtil;

import java.util.List;

/**
 * User:  huyunlong
 * Date:  2023/7/12
 */
public class Demo1 {
	public static void main(String[] args) {
//		DemoReport demoReport = new DemoReport();
//		demoReport.buildDemoDoc("\\C:\\Users\\huyunlong\\Desktop\\DIP运营分析报告-报告模板.docx",
//				"C:\\Users\\huyunlong\\Desktop\\郑州蓝博电子技术有限公司_DIP运营分析报告.docx");

		List<String> labels = PoiTlUtil.getLabel("\\\\C:\\\\Users\\\\huyunlong\\\\Desktop\\\\DIP运营分析报告-报告模板.docx");
		System.out.println(labels);

	}
}
