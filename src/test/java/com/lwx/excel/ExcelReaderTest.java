package com.lwx.excel;

import jdk.internal.util.xml.impl.Input;
import org.junit.Assert;
import org.junit.Test;

import java.io.InputStream;
import java.util.Arrays;

/**
 * @Author liuax01
 * @Date 2018/1/6 17:29
 */
public class ExcelReaderTest {

	@Test
	public void testXls总行数(){

		String fileName = "test1.xls";
		//System.out.println(ExcelReaderTest.class.getResource("/excel/"+fileName));
		InputStream is = ExcelReaderTest.class.getResourceAsStream("/excel/"+fileName);
		ExcelReader reader = new ExcelReader(is,fileName);
		Assert.assertEquals(7,reader.getTotalRow());
		String[][] data = reader.getData();
		printData(data);
	}

	@Test
	public void testXlsx总行数(){

		String fileName = "test1.xlsx";
		//System.out.println(ExcelReaderTest.class.getResource("/excel/"+fileName));
		InputStream is = ExcelReaderTest.class.getResourceAsStream("/excel/"+fileName);
		ExcelReader reader = new ExcelReader(is,fileName);
		Assert.assertEquals(12,reader.getTotalRow());
		String[][] data = reader.getData();
		printData(data);
	}

	private void printData(String[][] data){
		for (String[] data1 : data) {
			if (data1 == null) continue;
			for (String s : data1) {
				System.out.print(s+"\t");
			}
			System.out.println();
		}
	}
}
