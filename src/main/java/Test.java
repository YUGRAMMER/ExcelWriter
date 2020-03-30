
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.forcewin.excelWriter.ExcelWriter;
import com.forcewin.excelWriter.service.Service;
import com.forcewin.excelWriter.vo.ExcelFileVO;
import com.forcewin.excelWriter.vo.ExcelRowVO;
import com.forcewin.excelWriter.vo.ExcelSheetVO;


public class Test {

	public static void main(String[] args) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		// 탐색 대상 DB 구조 가져오기
		Connection conn = null;
		ResultSetMetaData rsmd = null;
		Statement stmt = null;
		
		try {
			// 입력 파라미터
			String fileName = "검출결과";
			int sample_data_count = 5;
			
			
			
			// 엑셀 헤더 스타일 설정
			CellStyle headerStyle = workbook.createCellStyle();
			headerStyle.setBorderRight(BorderStyle.THIN);
			headerStyle.setBorderLeft(BorderStyle.THIN);
			headerStyle.setBorderTop(BorderStyle.THIN);
			headerStyle.setBorderBottom(BorderStyle.THIN);
			headerStyle.setVerticalAlignment(VerticalAlignment.TOP);//높이 가운데 정렬
			headerStyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);  
			headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			// 헤더에 들어갈 문자열 배열
			List<String> headerValueList = new ArrayList<String>();
			String[] initHeaderValue = {"탐색 대상", "DB","테이블","필드","검출 항목","메타 정보","샘플 데이터"};
			for( String init_value : initHeaderValue ) headerValueList.add( init_value );
			for( int i = 1 ; i < sample_data_count; i ++ ) headerValueList.add("");

			// 엑셀 테이블 헤더 값 설정
			ExcelRowVO headerRow=new ExcelRowVO( headerValueList , headerStyle );
			
			// 셀 병합 객체
			CellRangeAddress merge1 = new CellRangeAddress(1,1,7,(7+sample_data_count)-1);

			// 입력 행 배열
			List<ExcelRowVO> rowList = new ArrayList<ExcelRowVO>();
			rowList.add(headerRow);


			Class.forName(Service.getClassName());
			conn =  DriverManager.getConnection(Service.getConnectionURL(),Service.getRepoID(),Service.getRepoPW());	
			

			// 셀통합 객체 배열
			List<CellRangeAddress> mergeList = new ArrayList<CellRangeAddress>();
			mergeList.add(merge1);

			ExcelSheetVO excelSheetVO = new ExcelSheetVO("DB_1",rowList,mergeList);
			ExcelFileVO excelFileVO = new ExcelFileVO(fileName,"xlsx","utf8",Arrays.asList(new ExcelSheetVO[]{excelSheetVO}));
			ExcelWriter ew = new ExcelWriter();
		
			ew.createExcel(workbook,excelFileVO);
		}catch(Exception e) {
			e.printStackTrace();
		}finally{
			if(stmt!=null)try{stmt.close();}catch(Exception e){e.printStackTrace();}
			if(conn!=null)try{conn.close();}catch(Exception e){e.printStackTrace();}
			try{workbook.close();}catch(Exception e){e.printStackTrace();}
			

		}
		
		
	}

}
