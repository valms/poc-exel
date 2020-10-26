package br.fjunior.poc.pocexel.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.Random;

@Component
public class ExcelUtils {
	
	
	public void createExcelFile() {
		
	
	}
	
	public static void main(String[] args) {
		try (
			SXSSFWorkbook xssfWorkbook = new SXSSFWorkbook();
			OutputStream outputStream = new FileOutputStream( "test.xlsx" )
		
		) {
			
			SXSSFSheet lineChartSheetTab = xssfWorkbook.createSheet( "lineChartSheet" );
			
			// Max rows supported by Exel: 1.048.576 rows
			for (int i = 0; i < 1048576; i++) {
				
				Cell cell = lineChartSheetTab
					            .createRow( i )
					            .createCell( 0 );
				
				cell.setCellValue( new BigDecimal( new Random().nextInt( 100000 ) ).intValue() );
			}
			
			int numOfColumns = lineChartSheetTab.getRow( 0 ).getPhysicalNumberOfCells();
			
			lineChartSheetTab.setAutoFilter( new CellRangeAddress( 0, 0, 0, 0 );
			
			xssfWorkbook.write( outputStream );
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private void populateRows() {
		new Random().nextInt( 100000 );
	}
	
}
