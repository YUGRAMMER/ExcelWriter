package com.forcewin.excelWriter.vo;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * POI 엑셀 셀 정보 객체
 */
public class ExcelCellVO {
    private CellStyle cellStyle;
    private String cellValue;
    public ExcelCellVO(){}
    public ExcelCellVO(String cellValue){
        this.cellValue=cellValue;
    }
    public ExcelCellVO(CellStyle cellStyle){
        this.cellStyle=cellStyle;
    }
    public ExcelCellVO(String cellValue,CellStyle cellStyle){
        this.cellValue=cellValue;
        this.cellStyle=cellStyle;
    }
    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public String getCellValue() {
        return cellValue;
    }

    public void setCellValue(String cellValue) {
        this.cellValue = cellValue;
    }
    
}