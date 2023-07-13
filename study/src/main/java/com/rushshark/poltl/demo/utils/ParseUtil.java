package com.rushshark.poltl.demo.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * User:  huyunlong
 * Date:  2023/7/12
 */
public class ParseUtil {
	public static final String REG_TL = "\\{\\{([^}]*)\\}\\}";

	public static void main(String[] args) {
		List<String> label = null;
		label = getLabel("C:\\Users\\huyunlong\\Desktop\\DIP运营分析报告-报告模板.docx");
		System.out.println(label);
	}
	/**
	 * 获取word文档中所有Poi_tl标签
	 * @param fileName
	 * @return
	 */
	public static List<String> getLabel(String fileName) {
		ArrayList<String> list = new ArrayList<>();
		File file = new File(fileName);
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(file);
			int readLine;
			byte[] data = new byte[4096];
			while ((readLine = fis.read(data)) != -1){
				String s = new String(data, 0, readLine);
				Pattern compile = Pattern.compile(REG_TL);
				Matcher matcher = compile.matcher(s);
				while (matcher.find()){
					String group = matcher.group();
					String label = group.substring(2, group.length() - 2);
					list.add(label);
				}
			}
		} catch (IOException e) {
			e.printStackTrace();
		}finally {
			try {
				fis.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		return list;
	}
}
