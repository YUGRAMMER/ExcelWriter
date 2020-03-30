package com.forcewin.excelWriter.vo;

import java.util.List;

public class ExcelRowVO {
	private List<String> cellList;
	public ExcelRowVO() {}
	public ExcelRowVO(List<String> cellList) {
		this.cellList=cellList;
	}
	public List<String> getCellList() {
		return cellList;
	}

	public void setCellList(List<String> cellList) {
		this.cellList = cellList;
	}
	
}
