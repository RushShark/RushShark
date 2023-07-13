package com.rushshark.poitl.doc;

import com.rushshark.poitl.demo.doc.DemoReport;
import com.rushshark.poitl.demo.model.DemoModelBuilder;
import com.rushshark.poitl.demo.utils.PoiTlUtil;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.List;

/**
 * User:  huyunlong
 * Date:  2023/7/13
 */
@SpringBootTest
public class DemoReportTest {
	@Autowired
	private DemoReport demoReport;
	@Autowired
	private DemoModelBuilder demoModelBuilder;

	@Test
	public void buildDemoDoc(){
		demoReport.buildDemoDoc("\\C:\\Users\\huyunlong\\Desktop\\DIP运营分析报告-报告模板.docx",
				"C:\\Users\\huyunlong\\Desktop\\郑州蓝博电子技术有限公司_DIP运营分析报告.docx",
				demoModelBuilder.buildData());
	}

	@Test
	public void getLabels(){
		List<String> label = PoiTlUtil.getLabel("\\C:\\Users\\huyunlong\\Desktop\\DIP运营分析报告-报告模板.docx");
		System.out.println(label);
	}
}
