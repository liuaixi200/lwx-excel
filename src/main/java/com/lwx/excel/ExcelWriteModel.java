package com.lwx.excel;

/**
 * @Author liuax01
 * @Date 2018/1/6 16:24
 */
public class ExcelWriteModel {

	private String value;
	private int width;
	private int height;
	private EnumExcelColor fontColor = EnumExcelColor.BLACK;
	private EnumExcelColor backGroundColor = EnumExcelColor.WHITE;
	private boolean boldWeight;

	public int getHeight() {
		return height*256;
	}

	public void setHeight(int height) {
		this.height = height;
	}

	public boolean isBoldWeight() {
		return boldWeight;
	}

	public void setBoldWeight(boolean boldWeight) {
		this.boldWeight = boldWeight;
	}

	public String getValue() {
		return value;
	}

	public void setValue(String value) {
		this.value = value;
	}

	public int getWidth() {
		return width*256;
	}

	public void setWidth(int width) {
		this.width = width;
	}

	public EnumExcelColor getFontColor() {
		return fontColor;
	}

	public void setFontColor(EnumExcelColor fontColor) {
		this.fontColor = fontColor;
	}

	public EnumExcelColor getBackGroundColor() {
		return backGroundColor;
	}

	public void setBackGroundColor(EnumExcelColor backGroundColor) {
		this.backGroundColor = backGroundColor;
	}
}
