package heathers.face;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.format.CellDateFormatter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelParse {
	private XSSFFormulaEvaluator evaluator;

	public ExcelParse(String filename){
		try {
			InputStream in = new FileInputStream(filename);
			XSSFWorkbook wbo = new XSSFWorkbook(in);
			XSSFSheet sheet = wbo.getSheetAt(35);
		
			evaluator = wbo.getCreationHelper().createFormulaEvaluator();
			
			Iterator<Row> rowIterator = sheet.iterator();
			Iterator<Cell> row;
			Cell cell;
			String value;
			DataFormatter df = new DataFormatter();
			while(rowIterator.hasNext()){
				row = rowIterator.next().cellIterator();
				while(row.hasNext()){
					cell = row.next();
					value = parseCell(cell);
					if(!value.equals("")){
						System.out.println(value.replaceAll("\n", ""));
					}
				}
			}
			in.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	public String parseCell(Cell cell){
		
		//If it's a formula and needs to be evaluated
		CellValue cellValue = evaluator.evaluate(cell);
		if(cellValue!=null){
			switch(cellValue.getCellType()){
				case Cell.CELL_TYPE_NUMERIC:
					double dv = cellValue.getNumberValue();
					if(DateUtil.isCellDateFormatted(cell)){
						Date date = DateUtil.getJavaDate(dv);
						String dateFormat = cell.getCellStyle().getDataFormatString();
						return new CellDateFormatter(dateFormat).format(date)+"";
					}else{
						return dv+"";
					}
				case Cell.CELL_TYPE_STRING:
					return cellValue.getStringValue();
			}
		}
		
		//If just a value
		switch(cell.getCellType()){
			case Cell.CELL_TYPE_NUMERIC:
				double dv = cell.getNumericCellValue();
				if(DateUtil.isCellDateFormatted(cell)){
					Date date = DateUtil.getJavaDate(dv);
					String dateFormat = cell.getCellStyle().getDataFormatString();
					return new CellDateFormatter(dateFormat).format(date)+"";
				}else{
					return dv+"";
				}
			case Cell.CELL_TYPE_STRING:
				return cell.getRichStringCellValue()+"";
			default: 
				cell.setCellType(Cell.CELL_TYPE_STRING);
				return cell.getStringCellValue()+"";
		}
	}
	
	public static void main(String[] args){
		new ExcelParse("src/schedules.xlsx");
	}
}
