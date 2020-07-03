package com.trs.txt2excel.parser;

import java.util.ArrayList;
import java.util.List;

public class TxtLineParser {

	public static List<String> readLine(String line) {
		List<String> items = new ArrayList<String>();
		StringBuilder builder = new StringBuilder();
		boolean noclose = false;
		for (char ch : line.toCharArray()) {
			if (ch == '"' || ch == '[' || ch == ']') {
				noclose = !noclose;
				continue;
			}
			if (noclose == false && ch == ' ') {
				items.add(builder.toString());
				builder = new StringBuilder();
				continue;
			}
			builder.append(ch);
		}
		items.add(builder.toString());
		return items;
	}
}
