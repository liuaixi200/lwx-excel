package com.lwx.excel;

import org.junit.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @Author liuax01
 * @Date 2018/1/6 17:52
 */
public class ExcelWriterTest {

	@Test
	public void testWriteXls() throws FileNotFoundException {
		FileOutputStream fos = new FileOutputStream(new File("D:\\data\\test/test.xls"));
		ExcelWriter writer = new ExcelWriter(getList(20000));
		writer.write2003(fos);
	}

	@Test
	public void testWriteXlsx() throws FileNotFoundException {
		FileOutputStream fos = new FileOutputStream(new File("D:\\data\\test/test.xlsx"));
		ExcelWriter writer = new ExcelWriter(getList(20000));
		writer.write2007(fos);
	}

	public List<ExcelWriteModel[]> getList(int num){
		List<ExcelWriteModel[]> list = new ArrayList<>();
		for(int i=0;i<num;i++){
			ExcelWriteModel[] ms = new ExcelWriteModel[10];
			for(int j=0;j<10;j++){
				ExcelWriteModel model = new ExcelWriteModel();
				model.setValue("中国人");
				ms[j]=model;
			}
			list.add(ms);
		}
		return list;
	}
}
