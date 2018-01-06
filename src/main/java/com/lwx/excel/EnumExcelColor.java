package com.lwx.excel;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Color;

/**
 * @Author liuax01
 * @Date 2018/1/6 17:04
 */
public enum  EnumExcelColor {

	RED(HSSFColor.RED.index),WHITE(HSSFColor.WHITE.index),BLACK(HSSFColor.BLACK.index),
	;

	private short color;

	private EnumExcelColor(short color){
		this.color = color;
	}

	public short getColor() {
		return color;
	}

	public void setColor(short color) {
		this.color = color;
	}
}
