package com.trs.txt2excel.util;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;

public class FileUtil {

	public static String loadFirstLine(File file, String charsetName) throws IOException {
		if (file == null || false == file.isFile()) {
			throw new IllegalArgumentException("the file [ " + file + " ] is NOT a file!");
		}
		try (InputStreamReader isr = new InputStreamReader(new FileInputStream(file), charsetName);
				BufferedReader br = new BufferedReader(isr);) {
			return br.readLine();
		} catch (IOException e) {
			throw e;
		}
	}

	public static List<String> loadText(File file, String encoding) throws IOException {
		assertFileCanRead(file);
		List<String> result = new ArrayList<String>();
		try (InputStreamReader isr = new InputStreamReader(new FileInputStream(file), encoding);
				BufferedReader br = new BufferedReader(isr);) {
			for (String line = br.readLine(); line != null; line = br.readLine()) {
				if (line.trim().isEmpty()) {
					continue;
				}
				result.add(line);
			}
			return result;
		} catch (IOException e) {
			throw e;
		}
	}

	/**
	 * 断言文件存在，否则抛出运行时异常.
	 * 
	 * @param f 文件对象
	 */
	public static void assertFileExists(File f) {
		if (f == null || false == f.isFile()) {
			throw new IllegalArgumentException("file not exist: [" + f + "]");
		}
	}

	/**
	 * 断言给定的文件可以被执行该JVM进程的用户读取(隐含该文件存在).
	 * 
	 * @param file 给定的文件
	 */
	public static void assertFileCanRead(File file) {
		assertFileExists(file);
		if (false == file.canRead()) {
			throw new IllegalArgumentException("the file [" + file + "] can not read");
		}
	}

}
