package com.artisan.util;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
	
	/**
	 * 
	 * @param sheet
	 * @param size
	 * @desc 自适应宽度 支持中文
	 */
	public static void autoSizeColumn(XSSFSheet sheet, int size) {

		// 使用autoSizeColumn方法可以把Excel设置为根据内容自动调整列宽，仅支持英文和数字,对中文无效
		for (int i = 0; i < size; i++) {
			sheet.autoSizeColumn(i);
		}
		// 中文自适应宽度的处理
		for (int columnNum = 0; columnNum < size; columnNum++) {
			int columnWidth = sheet.getColumnWidth(columnNum) / 256;
			for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
				XSSFRow currentRow;
				// 当前行未被使用过
				if (sheet.getRow(rowNum) == null) {
					currentRow = sheet.createRow(rowNum);
				} else {
					currentRow = sheet.getRow(rowNum);
				}

				if (currentRow.getCell(columnNum) != null) {
					XSSFCell currentCell = currentRow.getCell(columnNum);
					if (currentCell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
						int length = currentCell.getStringCellValue().getBytes().length;
						if (columnWidth < length) {
							columnWidth = length;
						}
					}
				}
			}
			sheet.setColumnWidth(columnNum, columnWidth * 256);
		}
	}
	
	
	/**
	 * 
	 * @param workbook
	 * @param bg 前景色
	 * @param hasFrame 是否设置四条边框
	 * @return
	 */
	public static CellStyle setStyle(XSSFWorkbook workbook, Short bg, boolean hasFrame) {
		CellStyle style = workbook.createCellStyle();
		// 设置颜色和边框
		if(null != bg ) {
			style.setFillForegroundColor(bg);
			style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		}
		// 四条边框
		if (hasFrame) {
			style.setBorderTop(XSSFCellStyle.BORDER_THIN);
			style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			style.setBorderRight(XSSFCellStyle.BORDER_THIN);
		}
		//设置自动换行
		//style.setWrapText(true);
		
		return style;
	}
	
	
	
}
