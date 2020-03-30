package com.forcewin.excelWriter.vo;

import java.util.List;

/**
 * Excel Data Object Used by ExcelWriter
 * @author dbwls
 *
 */
public class ExcelFileVO {
	private String fileName;
	private String fileExt;
	private String fileEncoding;
	private List<ExcelSheetVO> sheetList;
	
	public ExcelFileVO() {}
	public ExcelFileVO(String fileName, String fileExt, String fileEncoding, List<ExcelSheetVO> sheetList) {
		this.fileName=fileName;
		this.fileExt=fileExt;
		this.fileEncoding=fileEncoding;
		this.sheetList=sheetList;
	}
	public String getFileName() {
		return fileName;
	}

	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	public String getFileExt() {
		return fileExt;
	}

	public void setFileExt(String fileExt) {
		this.fileExt = fileExt;
	}

	public List<ExcelSheetVO> getSheetList() {
		return sheetList;
	}

	public void setSheetList(List<ExcelSheetVO> sheetList) {
		this.sheetList = sheetList;
	}

	public String getFileEncoding() {
		return fileEncoding;
	}

	public void setFileEncoding(String fileEncoding) {
		this.fileEncoding = fileEncoding;
	}
	
	
}
