package com.forcewin.excelWriter.vo;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;

public class ExcelRowVO {
	private List<ExcelCellVO> cellList;
	public ExcelRowVO() {}
	public ExcelRowVO(List<ExcelCellVO> cellList) {
		this.cellList=cellList;
	}
	public ExcelRowVO(List<String> cellValueList,CellStyle cellStyle) {
		List<ExcelCellVO> tmp = new ArrayList<ExcelCellVO>();
		for( String cellValue : cellValueList )tmp.add( new ExcelCellVO( cellValue,cellStyle ) );
		this.cellList=tmp;
	}
	public List<ExcelCellVO> getCellList() {
		return cellList;
	}

	public void setCellList(List<ExcelCellVO> cellList) {
		this.cellList = cellList;
	}
	
}
