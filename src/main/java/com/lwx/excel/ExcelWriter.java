package com.lwx.excel;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * @Author liuax01
 * @Date 2018/1/6 16:24
 */
public class ExcelWriter {

	Logger logger = LoggerFactory.getLogger(this.getClass());
	private List<ExcelWriteModel[]> models;
	private int currentSheet;
	private Workbook workbook;
	private String sheetName;
	private int rowHeight = 2*256;
	private int cellWidth = 10*256;

	public ExcelWriter(List<ExcelWriteModel[]> modleList){
		models = modleList;
	}

	public void write2003(OutputStream os){
		workbook = new HSSFWorkbook();
		_writeExcel();
		try {
			workbook.write(os);
		} catch (IOException e) {
			throw new RuntimeException("写入流失败",e);
		}

	}

	public void write2007(OutputStream os){
		workbook = new XSSFWorkbook();
		_writeExcel();
		try {
			workbook.write(os);
		} catch (IOException e) {
			throw new RuntimeException("写入流失败",e);
		}
	}

	public void setSheetName(String sheetName){
		this.sheetName = sheetName;
	}

	private void _writeExcel(){
		currentSheet++;
		Sheet sheet = null;
		if(null ==sheetName){
			sheet = workbook.createSheet();
		} else {
			sheet = workbook.createSheet(sheetName);
		}

		int totalRow = models.size();
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setBorderBottom(CellStyle.BORDER_THIN); //下边框
		cellStyle.setBorderLeft(CellStyle.BORDER_THIN);//左边框
		cellStyle.setBorderTop(CellStyle.BORDER_THIN);//上边框
		cellStyle.setBorderRight(CellStyle.BORDER_THIN);//右边框
		Font font = workbook.createFont();
		DataFormat format = workbook.createDataFormat();
		cellStyle.setDataFormat(format.getFormat("@"));
		for (int i = 0; i < totalRow; i++) {
			Row tr = sheet.createRow(i);
			ExcelWriteModel[] modelAar = models.get(i);
			int height = rowHeight;
			if(modelAar[0].getHeight() > 0){
				height = modelAar[0].getHeight();
			}
			tr.setHeight((short)height);
			for(int j=0;j<modelAar.length;j++){
				ExcelWriteModel model = modelAar[j];
				Cell td = tr.createCell(j);

				if(model.isBoldWeight()){
					font.setBoldweight(Font.BOLDWEIGHT_BOLD);
				}
				if(j == 0){
					int width = cellWidth;
					if(model.getWidth() > 0){
						width = model.getWidth();
					}
					sheet.setColumnWidth(j,width);
				}
				font.setColor(model.getFontColor().getColor());
				cellStyle.setFillForegroundColor(model.getBackGroundColor().getColor());
				cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);


				td.setCellValue(model.getValue());
				cellStyle.setFont(font);
				td.setCellStyle(cellStyle);
			}
		}

	}
}
