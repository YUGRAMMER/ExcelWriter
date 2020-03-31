
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.forcewin.excelWriter.ExcelWriter;
import com.forcewin.excelWriter.service.Service;
import com.forcewin.excelWriter.vo.ExcelCellVO;
import com.forcewin.excelWriter.vo.ExcelFileVO;
import com.forcewin.excelWriter.vo.ExcelRowVO;
import com.forcewin.excelWriter.vo.ExcelSheetVO;


public class Test {

	public static void main(String[] args) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		// 탐색 대상 DB 구조 가져오기
		Connection conn = null;
		DatabaseMetaData metadata = null;
		Statement stmt = null;
		ResultSet DBRS = null, TableRS = null, ColRS = null;
		try {
			// 입력 파라미터
			String fileName = "검출결과";
			String targetName = "샘플 저장소";
			String item_name = "암호화 알고리즘";
			int sample_data_count = 5;
			
			
			
			// 엑셀 헤더 스타일 설정
			CellStyle headerStyle = workbook.createCellStyle();
			headerStyle.setBorderRight(BorderStyle.THIN);
			headerStyle.setBorderLeft(BorderStyle.THIN);
			headerStyle.setBorderTop(BorderStyle.THIN);
			headerStyle.setBorderBottom(BorderStyle.THIN);
			headerStyle.setVerticalAlignment(VerticalAlignment.TOP);//높이 가운데 정렬
			headerStyle.setFillForegroundColor(HSSFColor.LIME.index);  
			headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			// 탐색 대상 스타일 설정
			CellStyle targetNameStyle = workbook.createCellStyle();
			targetNameStyle.setBorderRight(BorderStyle.THIN);
			targetNameStyle.setBorderLeft(BorderStyle.THIN);
			targetNameStyle.setBorderTop(BorderStyle.THIN);
			targetNameStyle.setBorderBottom(BorderStyle.THIN);
			targetNameStyle.setVerticalAlignment(VerticalAlignment.TOP);//높이 가운데 정렬
			targetNameStyle.setFillForegroundColor(HSSFColor.GREY_80_PERCENT.index);  
			targetNameStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			Font font = workbook.createFont();
			font.setColor(IndexedColors.WHITE.getIndex());
			font.setBold(true);
			targetNameStyle.setFont(font);
			

			// 탐색 데이터베이스 스타일 설정
			CellStyle DBNameStyle = workbook.createCellStyle();
			DBNameStyle.setBorderRight(BorderStyle.THIN);
			DBNameStyle.setBorderLeft(BorderStyle.THIN);
			DBNameStyle.setBorderTop(BorderStyle.THIN);
			DBNameStyle.setBorderBottom(BorderStyle.THIN);
			DBNameStyle.setVerticalAlignment(VerticalAlignment.TOP);//높이 가운데 정렬
			DBNameStyle.setFillForegroundColor(HSSFColor.GREY_50_PERCENT.index);  
			DBNameStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			font.setBold(false);
			DBNameStyle.setFont(font);

			
			// 탐색 테이블 스타일 설정
			CellStyle TableNameStyle = workbook.createCellStyle();
			TableNameStyle.setBorderRight(BorderStyle.THIN);
			TableNameStyle.setBorderLeft(BorderStyle.THIN);
			TableNameStyle.setBorderTop(BorderStyle.THIN);
			TableNameStyle.setBorderBottom(BorderStyle.THIN);
			TableNameStyle.setVerticalAlignment(VerticalAlignment.TOP);//높이 가운데 정렬
			TableNameStyle.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);  
			TableNameStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			// 탐색 필드 스타일 설정
			CellStyle FieldNameStyle = workbook.createCellStyle();
			FieldNameStyle.setBorderRight(BorderStyle.THIN);
			FieldNameStyle.setBorderLeft(BorderStyle.THIN);
			FieldNameStyle.setBorderTop(BorderStyle.THIN);
			FieldNameStyle.setBorderBottom(BorderStyle.THIN);
			FieldNameStyle.setVerticalAlignment(VerticalAlignment.TOP);//높이 가운데 정렬
			FieldNameStyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);  
			FieldNameStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			// 일반 필드 스타일 설정
			CellStyle NormalStyle = workbook.createCellStyle();
			NormalStyle.setBorderRight(BorderStyle.THIN);
			NormalStyle.setBorderLeft(BorderStyle.THIN);
			NormalStyle.setBorderTop(BorderStyle.THIN);
			NormalStyle.setBorderBottom(BorderStyle.THIN);
			NormalStyle.setVerticalAlignment(VerticalAlignment.TOP);//높이 가운데 정렬

			// 일반 필드 스타일 설정
			CellStyle highlightStyle = workbook.createCellStyle();
			highlightStyle.setBorderRight(BorderStyle.THIN);
			highlightStyle.setBorderLeft(BorderStyle.THIN);
			highlightStyle.setBorderTop(BorderStyle.THIN);
			highlightStyle.setBorderBottom(BorderStyle.THIN);
			highlightStyle.setVerticalAlignment(VerticalAlignment.TOP);//높이 가운데 정렬
			highlightStyle.setFillForegroundColor(HSSFColor.GOLD.index);  
			highlightStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			// 헤더에 들어갈 문자열 배열
			List<String> headerValueList = new ArrayList<String>();
			String[] initHeaderValue = {"탐색 대상", "DB","테이블","필드","검출 항목","메타 정보","샘플 데이터"};
			for( String init_value : initHeaderValue ) headerValueList.add( init_value );
			for( int i = 1 ; i < sample_data_count; i ++ ) headerValueList.add("");

			// 엑셀 테이블 헤더 값 설정
			ExcelRowVO headerRow=new ExcelRowVO( headerValueList , headerStyle );
			
			
			// 입력 행 배열
			List<ExcelRowVO> rowList = new ArrayList<ExcelRowVO>();
			rowList.add(headerRow);
			// 셀통합 객체 배열
			List<CellRangeAddress> mergeList = new ArrayList<CellRangeAddress>();

			Class.forName(Service.getClassName());
			String ConnectionURL = Service.getConnectionURL();
			if( Service.getRepoName().equals("sybase")) ConnectionURL += "&allowMultiQueries=true";
			conn =  DriverManager.getConnection(ConnectionURL,Service.getRepoID(),Service.getRepoPW());	
			metadata = conn.getMetaData();
			DBRS =  Service.isSpecialRepo( Service.getRepoName())? metadata.getSchemas() : metadata.getCatalogs();
			// 셀 병합을 위한 인덱싱 변수
			int rowIndex = 2;
			int DBStartRowIndex = 2;
			int TableStartRowIndex = 2;
			Map<String,String> metadataMapper = Service.getMetaDataMapper();
			L1: while (DBRS.next()) {
				String db_name = DBRS.getString(1);
				if (Service.getDBFilter().get(db_name) != null) continue L1;
				String[] skip = { "TABLE" }, both = { "TABLE", "VIEW" }, only = { "VIEW" };

				if( Service.getViewoption().equals("skip") ){
					TableRS = Service.isSpecialRepo(Service.getRepoName()) ? metadata.getTables(null, db_name, "%", skip) : metadata.getTables(db_name, null, "%", skip);
				}else if( Service.getViewoption().equals("both") ){
					TableRS = Service.isSpecialRepo(Service.getRepoName()) ? metadata.getTables(null, db_name, "%", both) : metadata.getTables(db_name, null, "%", both);
				}else if( Service.getViewoption().equals("only") ){
					TableRS = Service.isSpecialRepo(Service.getRepoName()) ? metadata.getTables(null, db_name, "%", only): metadata.getTables(db_name, null, "%", only);
				}
				L2:while(TableRS.next()){
					String tb_name = TableRS.getString(3);
					if (Service.getTableFilter().get(tb_name) != null)continue L2;
					
					ColRS = Service.isSpecialRepo(Service.getRepoName()) ? metadata.getColumns(null, db_name, tb_name, null)
							: metadata.getColumns(db_name, null, tb_name, null);
					while (ColRS.next()) {
						String fld_name = ColRS.getString(4);
						rowIndex+=1;
						String report_meta = metadataMapper.get( String.format("%s,%s,%s", db_name,tb_name,fld_name ) );
						
						if( report_meta == null ){
							List<ExcelCellVO> cellList = new ArrayList<ExcelCellVO>();
							cellList.add(new ExcelCellVO(targetName,targetNameStyle));
							cellList.add(new ExcelCellVO(db_name,DBNameStyle));
							cellList.add(new ExcelCellVO(tb_name,TableNameStyle));
							cellList.add(new ExcelCellVO(fld_name,FieldNameStyle));
							cellList.add(new ExcelCellVO("",NormalStyle));
							cellList.add(new ExcelCellVO("",NormalStyle));
							for( int i = 0 ; i < sample_data_count; i ++ ) cellList.add(new ExcelCellVO("",NormalStyle));
							ExcelRowVO row = new ExcelRowVO( cellList );
							rowList.add(row);	
						}else{
							// 매칭 값이 있는 컬럼
							List<ExcelCellVO> cellList = new ArrayList<ExcelCellVO>();
							cellList.add(new ExcelCellVO(targetName,targetNameStyle));
							cellList.add(new ExcelCellVO(db_name,DBNameStyle));
							cellList.add(new ExcelCellVO(tb_name,TableNameStyle));
							cellList.add(new ExcelCellVO(fld_name,highlightStyle));
							cellList.add(new ExcelCellVO(item_name,highlightStyle));
							cellList.add(new ExcelCellVO(report_meta,highlightStyle));

							// 샘플 데이터 가져오기
							String repoName = Service.getRepoName();
							String query = String.format( "SELECT %s FROM %s.%s", fld_name,db_name,tb_name );
							
							if( repoName.equals("altibase") || repoName.equals("cubrid") ){
								query = String.format( "SELECT %s FROM %s.%s LIMIT 0 %d", fld_name,db_name,tb_name, sample_data_count );
							}if( repoName.equals("mysql") || repoName.equals("mariadb") || repoName.equals("h2") || repoName.equals("postgresql")){
								query = String.format( "SELECT %s FROM %s.%s LIMIT %d", fld_name,db_name,tb_name, sample_data_count );
							}else if( repoName.equals("sybase") ){
								query = String.format("SET ROWCOUNT %d;SELECT %s FROM %s.%s;",sample_data_count,db_name,tb_name,fld_name);
							}else if( repoName.equals("mssql") ){
								query = String.format( "SELECT top %d %s FROM %s.%s", sample_data_count, fld_name,db_name,tb_name );
							}else if( repoName.equals("oracle") ){
								query = String.format( "SELECT %s FROM %s.%s WHERE rownum<=%d", fld_name,db_name,tb_name, sample_data_count );
							}
							PreparedStatement pstmt = null;
							ResultSet tmpRS = null;
							int tmp_cnt = 0;
							try{
								pstmt = conn.prepareStatement(query);
								tmpRS = pstmt.executeQuery();
								
								L4:for( int i = 0 ; i < sample_data_count; i ++ ){
									
									if(tmpRS.isLast()){
										cellList.add(new ExcelCellVO("",highlightStyle));
										continue L4;
									}else{
										tmpRS.next();
										String sampleData = tmpRS.getString(fld_name);
										cellList.add(new ExcelCellVO(sampleData,highlightStyle));
									}
									
									tmp_cnt+=1;
								}

							}catch(Exception e){
								
								e.printStackTrace();
							}finally{
								if( tmpRS != null ) tmpRS.close();
								if( pstmt != null ) pstmt.close();
							}
							if(tmp_cnt==0)for( int i = 0 ; i < sample_data_count; i ++ ) cellList.add(new ExcelCellVO("",highlightStyle));
							
							ExcelRowVO row = new ExcelRowVO( cellList );
							rowList.add(row);	
						}
						
						
					}

					if( rowIndex != TableStartRowIndex ){
						CellRangeAddress merge_tb = new CellRangeAddress(TableStartRowIndex,rowIndex-1,3,3);
						mergeList.add(merge_tb);
					}else{
						// 컬럼이 없는 테이블
						List<ExcelCellVO> cellList = new ArrayList<ExcelCellVO>();
						cellList.add(new ExcelCellVO(targetName,targetNameStyle));
						cellList.add(new ExcelCellVO(db_name,DBNameStyle));
						cellList.add(new ExcelCellVO(tb_name,TableNameStyle));
						cellList.add(new ExcelCellVO("",FieldNameStyle));
						cellList.add(new ExcelCellVO("",NormalStyle));
						cellList.add(new ExcelCellVO("",NormalStyle));

						for( int i = 0 ; i < sample_data_count; i ++ ) cellList.add(new ExcelCellVO("",NormalStyle));
						ExcelRowVO row = new ExcelRowVO( cellList );
						rowList.add(row);	
						rowIndex+=1;
					}
					TableStartRowIndex=rowIndex;
				}
				if( rowIndex != DBStartRowIndex ){
					CellRangeAddress merge_db = new CellRangeAddress(DBStartRowIndex,rowIndex-1,2,2);
					mergeList.add(merge_db);
				}else{
					// 테이블이 없는 db
					List<ExcelCellVO> cellList = new ArrayList<ExcelCellVO>();
					cellList.add(new ExcelCellVO(targetName,targetNameStyle));
					cellList.add(new ExcelCellVO(db_name,DBNameStyle));
					cellList.add(new ExcelCellVO("",TableNameStyle));
					cellList.add(new ExcelCellVO("",FieldNameStyle));
					cellList.add(new ExcelCellVO("",NormalStyle));
					cellList.add(new ExcelCellVO("",NormalStyle));
					for( int i = 0 ; i < sample_data_count; i ++ ) cellList.add(new ExcelCellVO("",NormalStyle));
					ExcelRowVO row = new ExcelRowVO( cellList );
					rowList.add(row);	
					rowIndex+=1;
				}
				DBStartRowIndex=rowIndex;
				
				// 들어갈 문자열 배열
				
			}

			
			// 셀 병합 객체
			CellRangeAddress merge1 = new CellRangeAddress(1,1,7,(7+sample_data_count)-1);
			// 탐색 대상 영역 통합
			CellRangeAddress merge2 = new CellRangeAddress(2,rowIndex-1,1,1);
			mergeList.add(merge1);
			mergeList.add(merge2);

			ExcelSheetVO excelSheetVO = new ExcelSheetVO("DB_1",rowList,mergeList);
			ExcelFileVO excelFileVO = new ExcelFileVO(fileName,"xlsx","utf8",Arrays.asList(new ExcelSheetVO[]{excelSheetVO}));
			ExcelWriter ew = new ExcelWriter();
		
			ew.createExcel(workbook,excelFileVO);
		}catch(Exception e) {
			e.printStackTrace();
		}finally{
			if(ColRS!=null)try{ColRS.close();}catch(Exception e){e.printStackTrace();}
			if(TableRS!=null)try{TableRS.close();}catch(Exception e){e.printStackTrace();}
			if(DBRS!=null)try{DBRS.close();}catch(Exception e){e.printStackTrace();}
			if(stmt!=null)try{stmt.close();}catch(Exception e){e.printStackTrace();}
			if(conn!=null)try{conn.close();}catch(Exception e){e.printStackTrace();}
			try{workbook.close();}catch(Exception e){e.printStackTrace();}
			

		}
		
		
	}

}
