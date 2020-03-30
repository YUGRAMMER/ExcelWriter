package com.forcewin.excelWriter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.io.BufferedWriter;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.forcewin.excelWriter.vo.ExcelFileVO;
import com.forcewin.excelWriter.vo.ExcelRowVO;
import com.forcewin.excelWriter.vo.ExcelSheetVO;
/**
 * Simple create excel file Program.
 * Using Apache POI Library
 * @author dbwls
 */
public class ExcelWriter {
	/**
	 * create XSSFWorkbook( POI Excel object ) and return it.
	 * @param excelfileVO
	 * @return workbook
	 * @throws Exception
	 */
	public void createExcel(ExcelFileVO excelfileVO)throws Exception{
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		List<ExcelSheetVO> sheetList = excelfileVO.getSheetList();
		
		if( sheetList == null ) {
			if(workbook!=null)workbook.close();
			throw new Exception("No Sheet at Input VO.");
		}
		
		L1:for( ExcelSheetVO excelSheet : sheetList ) {
			XSSFSheet sheet = workbook.createSheet(excelSheet.getSheetName());
			List<CellRangeAddress> mergetList = excelSheet.getMergeList();
			
			List<ExcelRowVO> excelRowList = excelSheet.getRowList();
			if( excelRowList == null ) continue L1;
			
			int rowCount = 0;
			L2:for( ExcelRowVO excelRow : excelRowList ) {
				Row row = sheet.createRow(++rowCount);
				 int columnCount = 0;
				 List<String> cellList = excelRow.getCellList();
				//1.셀 스타일 및 폰트 설정
				 CellStyle headerStyle = workbook.createCellStyle();
				 headerStyle.setBorderRight(BorderStyle.THIN);
				 headerStyle.setBorderLeft(BorderStyle.THIN);
				 headerStyle.setBorderTop(BorderStyle.THIN);
				 headerStyle.setBorderBottom(BorderStyle.THIN);
				 headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);//높이 가운데 정렬


				 if( cellList == null ) continue L2;
				 for( String cellString : cellList ) {
					 Cell cell = row.createCell(++columnCount);
					 cell.setCellValue(cellString);
					 cell.setCellStyle(headerStyle);
	      
				 }
			}
			//셀병합
			if( mergetList!=null ) for( CellRangeAddress ca :mergetList )sheet.addMergedRegion(ca);
			

		}
		File tmpFile = new File(String.format("%s.%s", excelfileVO.getFileName(),excelfileVO.getFileExt() ));
		int cnt = 0;
		L1:while(true) {
			if( tmpFile.exists() ) {
				tmpFile = new File(String.format("%s_%d.%s", excelfileVO.getFileName(),++cnt,excelfileVO.getFileExt() ));
			}else {
				break L1;
			}
		}
		
		FileOutputStream fis = null;
		OutputStreamWriter osw = null;
		BufferedWriter bw = null;
		try {
			fis = new FileOutputStream( tmpFile.getAbsolutePath() );
			workbook.write(fis);
		}catch(Exception e) {
			throw e;
		}finally {
			if(bw!=null)bw.close();
			if(osw!=null)osw.close();
			if(fis!=null)fis.close();
			if(workbook!=null)workbook.close();
		}
		
	}
}
