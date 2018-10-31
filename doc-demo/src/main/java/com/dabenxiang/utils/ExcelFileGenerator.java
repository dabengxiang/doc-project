package com.dabenxiang.utils;

import com.dabenxiang.properties.MergeProperties;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.List;

public class ExcelFileGenerator {

	private final int SPLIT_COUNT = 50000; // Excel每个工作表的行数

	private List fieldName = null; // excel数据的抬头栏，即名称栏

	private List fieldData = null; // excel导出的实际数据

	private HSSFWorkbook workBook = null;// 一个excel文件

	private  int columnWidth  = 10000;

	private int rowHeight = 30;




	public ExcelFileGenerator(List titleCols, List<List<String>> exportDatas) {
		this.fieldName = titleCols;
		this.fieldData = exportDatas;
	}

	public HSSFWorkbook createWorkbook() {
		workBook = new HSSFWorkbook();
		int rows = fieldData.size();
		int sheetNum = 0;
		if (rows % SPLIT_COUNT == 0) {
			sheetNum = rows / SPLIT_COUNT;
		} else {
			sheetNum = rows / SPLIT_COUNT + 1;
		}

		HSSFCellStyle style = workBook.createCellStyle();
		// 设置这些样式
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
		style.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setAlignment(HorizontalAlignment.CENTER);
		// 生成一个字体
		HSSFFont font = workBook.createFont();
		font.setColor(IndexedColors.BLACK.getIndex());
		font.setFontHeightInPoints((short) 10);
		font.setBold(true);
		// font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		// 把字体应用到当前的样式
		style.setFont(font);

		// 生成并设置另一个样式
		HSSFCellStyle style2 = workBook.createCellStyle();
		style2.setBorderBottom(BorderStyle.THIN);
		style2.setBorderLeft(BorderStyle.THIN);
		style2.setBorderRight(BorderStyle.THIN);
		style2.setBorderTop(BorderStyle.THIN);
		style2.setAlignment(HorizontalAlignment.CENTER);
		style2.setVerticalAlignment(VerticalAlignment.CENTER);

		// 生成并设置另一个样式
		HSSFCellStyle style3 = workBook.createCellStyle();
		style3.setBorderBottom(BorderStyle.THIN);
		style3.setBorderLeft(BorderStyle.THIN);
		style3.setBorderRight(BorderStyle.THIN);
		style3.setBorderTop(BorderStyle.THIN);
		style3.setAlignment(HorizontalAlignment.CENTER);
		style3.setVerticalAlignment(VerticalAlignment.CENTER);

		for (int i = 1; i <= sheetNum; i++) {
			HSSFSheet sheet = workBook.createSheet("Page " + i);
			HSSFRow headRow = sheet.createRow(0);
			for (int j = 0; j < fieldName.size(); j++) {
				HSSFCell cell = headRow.createCell(j);
				// 设置单元格格式
				cell.setCellType(CellType.STRING);
//				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				sheet.setColumnWidth(j, 6000);
				// 将数据填入单元格
				if (fieldName.get(j) != null) {
					cell.setCellStyle(style);
					cell.setCellValue((String) fieldName.get(j));
				} else {
					cell.setCellStyle(style);
					cell.setCellValue("-");
				}
			}
			// 创建数据栏单元格并填入数据
			for (int k = 0; k < (rows < SPLIT_COUNT ? rows : SPLIT_COUNT); k++) {
				if (((i - 1) * SPLIT_COUNT + k) >= rows)
					break;
				HSSFRow row = sheet.createRow(k + 1);
				row.setHeightInPoints(rowHeight);

				ArrayList rowList = (ArrayList) fieldData.get((i - 1)
						* SPLIT_COUNT + k);
				for (int n = 0; n < rowList.size(); n++) {
					HSSFCell cell = row.createCell(n);
					if (rowList.get(n) != null) {

						if (n !=1)
							cell.setCellStyle(style2);
						else
							cell.setCellStyle(style3);
						HSSFRichTextString richString = new HSSFRichTextString(
								(String) rowList.get(n));
						cell.setCellValue(richString);
					} else {
						cell.setCellValue("");
					}
				}
			}
		}
		return workBook;
	}

	public void exportExcel(OutputStream os) throws Exception {
		workBook = createWorkbook();
		workBook.write(os);
		os.close();
	}


	public void exportExcel(OutputStream os, List<MergeProperties> mergeList) throws Exception {
		workBook = createWorkbook();

		HSSFSheet sheetAt = workBook.getSheetAt(0);

		if(mergeList!=null ){
			for (MergeProperties mergeProperties : mergeList) {
				CellRangeAddress cra1 =new CellRangeAddress(mergeProperties.getStartRow(),
						mergeProperties.getEndRow(),
						mergeProperties.getStartCol(),
						mergeProperties.getEndCol()); // 起始行, 终止行, 起始列, 终止列
				sheetAt.addMergedRegion(cra1);
			}

		}



		workBook.write(os);
		os.close();
	}






	public  static CellStyle getDefaultCellStyle(Workbook wb){
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyle.setWrapText(true);
		return cellStyle;


	}


	public static HSSFWorkbook createExcel(OutputStream outputStream, String[][] table) {

		Workbook wb = new HSSFWorkbook();
		Font titleFont = wb.createFont();
		CellStyle titleStyle = getDefaultCellStyle(wb);
		CellStyle valueCellStyle = getDefaultCellStyle(wb);
		CellStyle keyCellStyle = getDefaultCellStyle(wb);

		titleStyle.setAlignment(HorizontalAlignment.CENTER);
		titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		keyCellStyle.setAlignment(HorizontalAlignment.CENTER);
		keyCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);




		titleFont.setBold(true);
		titleFont.setFontHeightInPoints((short) 14);
		titleStyle.setFont(titleFont);

		Sheet sheet =  wb.createSheet();


		sheet.setDisplayGridlines(false);

		sheet.setColumnWidth(0,14 * 256);
		sheet.setColumnWidth(1,50 * 256);
		sheet.setColumnWidth(2,18 * 256);
		sheet.setColumnWidth(3,50 * 256);


//        CellRangeAddress cra1 =new CellRangeAddress(0, 0, 0, 3); // 起始行, 终止行, 起始列, 终止列
//        CellRangeAddress cra2 =new CellRangeAddress(3, 3, 1, 3); // 起始行, 终止行, 起始列, 终止列
//        CellRangeAddress cra3 =new CellRangeAddress(8, 8, 1, 3); // 起始行, 终止行, 起始列, 终止列
//
//
//        sheet.addMergedRegion(cra1);
//        sheet.addMergedRegion(cra2);
//        sheet.addMergedRegion(cra3);


		for (int i = 0; i < table.length; i++) {
			Row row = sheet.createRow(i);
			if(i==0){
				row.setHeightInPoints(50);
			}else {
				row.setHeightInPoints(40);

			}

			boolean flag = false;

			for (int j = 0; j < 4; j++) {
				Cell cell = row.createCell(j);
				if(j%2 == 0)
					cell.setCellStyle(keyCellStyle);
				else
					cell.setCellStyle(valueCellStyle);
				if(table[i][j] == null){
					if(flag)
						continue;
					else{
						if(j==0){
							row.removeCell(cell);
							break;

						}
						CellRangeAddress cra =new CellRangeAddress(i, i, j-1, table[i].length-1); // 起始行, 终止行, 起始列, 终止列
						sheet.addMergedRegion(cra);
						flag = true;
					}
					continue;
				}
				if(i==0){
					cell.setCellStyle(titleStyle);
					cell.setCellValue(table[0][0]);
					continue;
				}
				cell.setCellValue(table[i][j]);

			}
		}
//		Iterator<Row> rows = sheet.rowIterator();
//
//		while(rows.hasNext()){
//			Row row = (Row) rows.next();
//			if(row!=null) {
//				int num = row.getLastCellNum();
//				System.out.println(num);
//			}
//		}
//
//
//		FileOutputStream fileOut = null;
//		try {
//			wb.write(outputStream);
//			fileOut.close();
//		} catch (Exception e) {
//			e.printStackTrace();
//		}


		return (HSSFWorkbook) wb;

	}


	public static void main(String[] args) throws Exception {
		List<String> titleList = new ArrayList<>();
		List<List<String>> columList = new ArrayList<>();

		titleList.add("序号");
		titleList.add("标题");
		titleList.add("数量");

		for (int i = 0; i < 5; i++) {
			List<String> valueList = new ArrayList<>();
			valueList.add("2");
			valueList.add("22222222222222222222222222222222222");
			valueList.add("2");
			columList.add(valueList);

		}

		ExcelFileGenerator excelFileGenerator = new ExcelFileGenerator(titleList,columList);
		String targetFilePath =  "C:\\Users\\dabenxiang\\Desktop\\1.xls";

		FileOutputStream fileOutputStream = new FileOutputStream(targetFilePath);


		List<MergeProperties> mergeList = new ArrayList<>();

		MergeProperties mergeProperties = new MergeProperties(2, 3, 1, 1);


		mergeList.add(mergeProperties);


		excelFileGenerator.exportExcel(fileOutputStream,mergeList);

	}



}