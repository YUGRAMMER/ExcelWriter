import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.util.CellRangeAddress;

import com.forcewin.excelWriter.ExcelWriter;
import com.forcewin.excelWriter.vo.ExcelFileVO;
import com.forcewin.excelWriter.vo.ExcelRowVO;
import com.forcewin.excelWriter.vo.ExcelSheetVO;

public class main {

	public static void main(String[] args) {
		// header row
		ExcelRowVO hr = new ExcelRowVO(Arrays.asList(new String[]{ "탐색대상", "DB", "테이블", "필드" , "검출 항목","메타 정보","샘플 데이터"}));

		ExcelRowVO cr1 = new ExcelRowVO(Arrays.asList(new String[]{ "첫번째 저장소", "DB1", "tb_1", "fld1_1_1" , "암호화 플러그인","HASH","fa39f64aa3510b604de9328c59383f6457a35197","3858f62230ac3c915f300c664312c63f","09234807e4af85f17c66b48ee3bca89dffd1f1233659f9f940a2b17b0b8c6bc5"}));
		ExcelRowVO cr2 = new ExcelRowVO(Arrays.asList(new String[]{ "", "", "", "fld1_1_2","","","","",""}));
		ExcelRowVO cr3 = new ExcelRowVO(Arrays.asList(new String[]{ "", "", "", "fld1_1_3" , "암호화 플러그인","DES or 3DES? (0)","fa39f64aa3510b604de9328c59383f6457a35197","3858f62230ac3c915f300c664312c63f","09234807e4af85f17c66b48ee3bca89dffd1f1233659f9f940a2b17b0b8c6bc5"}));
		ExcelRowVO cr4 = new ExcelRowVO(Arrays.asList(new String[]{ "", "", "tb_2", "fld1_2_1" , "암호화 플러그인","HASH","fa39f64aa3510b604de9328c59383f6457a35197","3858f62230ac3c915f300c664312c63f","09234807e4af85f17c66b48ee3bca89dffd1f1233659f9f940a2b17b0b8c6bc5"}));
		ExcelRowVO cr5 = new ExcelRowVO(Arrays.asList(new String[]{ "", "", "tb_2", "fld1_2_2","","","","",""}));
		ExcelRowVO cr6 = new ExcelRowVO(Arrays.asList(new String[]{ "", "", "tb_2", "fld1_2_3","","","","",""}));
		ExcelRowVO cr7 = new ExcelRowVO(Arrays.asList(new String[]{ "", "DB2", "tb_1", "fld2_1_1","암호화 플러그인","DES or 3DES? (0)","fa39f64aa3510b604de9328c59383f6457a35197","3858f62230ac3c915f300c664312c63f","09234807e4af85f17c66b48ee3bca89dffd1f1233659f9f940a2b17b0b8c6bc5"}));
		ExcelRowVO cr8 = new ExcelRowVO(Arrays.asList(new String[]{ "", "", "", "fld2_1_2","암호화 플러그인","SHA2-384 (384)","fa39f64aa3510b604de9328c59383f6457a35197","3858f62230ac3c915f300c664312c63f","09234807e4af85f17c66b48ee3bca89dffd1f1233659f9f940a2b17b0b8c6bc5"}));
		ExcelRowVO cr9 = new ExcelRowVO(Arrays.asList(new String[]{ "", "", "tb2", "fld2_2_1","","","","",""}));
		
		CellRangeAddress merge1 = new CellRangeAddress(2,10,1,1);
		ExcelSheetVO excelSheetVO = new ExcelSheetVO("DB_1",Arrays.asList(new ExcelRowVO[] {hr,cr1,cr2,cr3,cr4,cr5,cr6,cr7,cr8,cr9}),Arrays.asList(new CellRangeAddress[] {merge1}));
		ExcelFileVO excelFileVO = new ExcelFileVO("검출결과","xlsx","utf8",Arrays.asList(new ExcelSheetVO[]{excelSheetVO}));
		ExcelWriter ew = new ExcelWriter();
		try {
			ew.createExcel(excelFileVO);
		}catch(Exception e) {
			e.printStackTrace();
		}
		
		
	}

}
