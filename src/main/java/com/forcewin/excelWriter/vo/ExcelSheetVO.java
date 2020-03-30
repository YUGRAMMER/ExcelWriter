package com.forcewin.excelWriter.vo;

import java.util.List;

import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelSheetVO {
	private String sheetName;
	private List<ExcelRowVO> rowList;
	private List<CellRangeAddress> mergeList;
	public ExcelSheetVO() {}
	public ExcelSheetVO(String sheetName, List<ExcelRowVO> rowList,List<CellRangeAddress> mergeList) {
		this.sheetName=sheetName;
		this.rowList=rowList;
		this.mergeList=mergeList;
	}
	public String getSheetName() {
		return sheetName;
	}
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	public List<ExcelRowVO> getRowList() {
		return rowList;
	}
	public void setRowList(List<ExcelRowVO> rowList) {
		this.rowList = rowList;
	}
	public List<CellRangeAddress> getMergeList() {
		return mergeList;
	}
	public void setMergeList(List<CellRangeAddress> mergeList) {
		this.mergeList = mergeList;
	}
	
}
