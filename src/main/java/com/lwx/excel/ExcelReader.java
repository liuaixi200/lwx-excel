package com.lwx.excel;

import ch.qos.logback.classic.LoggerContext;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;

/**
 * @Author liuax01
 * @Date 2018/1/6 13:42
 */
public class ExcelReader {

	Logger logger = LoggerFactory.getLogger(this.getClass());

	private int totalRow;

	private final String EXCEL_V_2003 = ".xls";
	private final String EXCEL_V_2007 = ".xlsx";
	private Workbook workbook;
	private Sheet sheet;
	private int currentSheet=0;//默认解析第1页
	private int maxCollunms;
	private String[][] data;

	public ExcelReader(String fileName){
		if(null == fileName || fileName.length() == 0){
			throw new RuntimeException("文件名不能为空");
		}
		File file = new File(fileName);
		FileInputStream fis;
		try {
			fis = new FileInputStream(file);
		} catch (FileNotFoundException e) {
			throw new RuntimeException("读取文件失败",e);
		}
		if(file.isDirectory()){
			throw new RuntimeException("不能是目录");
		}

		_createWorkbook(fis,fileName);
	}

	public ExcelReader(InputStream is,String fileName){
		_createWorkbook(is,fileName);
	}

	private void _createWorkbook(InputStream is,String fileName){
		if(fileName.endsWith(EXCEL_V_2003)){
			try {
				workbook = new HSSFWorkbook(is);
			} catch (IOException e) {
				throw new RuntimeException("读取EXCEL失败",e);
			}
		} else if(fileName.endsWith(EXCEL_V_2007)){
			try {
				workbook = new XSSFWorkbook(is);
			} catch (IOException e) {
				throw new RuntimeException("读取EXCEL失败",e);
			}
		} else {
			throw new RuntimeException("无法识别的扩展名");
		}
		_parseExcel();
	}

	private void _parseExcel(){
		logger.info(String.format("当前解析第%d页",currentSheet));
		sheet = workbook.getSheetAt(currentSheet);
		totalRow = sheet.getLastRowNum()+1;
		data = new String[totalRow][];
		for(int i=0;i<totalRow;i++){
			Row tr = sheet.getRow(i);
			if(null == tr){
				continue;
			}
			int totalCells = tr.getLastCellNum()+1;
			if(totalCells>maxCollunms){
				maxCollunms = totalCells;
			}
			String[] tds = new String[totalCells];
			data[i] = tds;
			for(int j = 0;j<totalCells;j++){
				Cell td = tr.getCell(j);
				String value = getValue(td);
				tds[j] = value;
			}
		}
	}

	public String[][] getData(){
		return data;
	}

	public void parseExcel(int book){
		this.currentSheet = book;
		_parseExcel();
	}

	private String getValue(Cell cell){
		String cellStr = "";
		if (cell == null) {// 单元格为空设置cellStr为空串
			cellStr = "";
		} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {// 对布尔值的处理
			cellStr = String.valueOf(cell.getBooleanCellValue());
		} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {// 对数字值的处理
			cellStr = (int)cell.getNumericCellValue() + "";
		} else {// 其余按照字符串处理
			cellStr = cell.getStringCellValue();
		}
		return cellStr;
	}

	public int getTotalRow(){
		return totalRow;
	}
}
