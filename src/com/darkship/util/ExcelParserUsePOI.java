package com.darkship.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import test.TestExcel;

public class ExcelParserUsePOI {
	
	private ExcelParserUsePOI(){};
	
	/**
	 * @Method : parseRegularExcel
	 * @version : 0.11
	 * @Date : 2012. 9. 13.
	 * @author : kim munsu
	 * 정돈된 excel을 parsing후, 세번째 인자로 받는 Class 를 객체 생성하여 list에 담아 return합니다. 
	 * 
	 * @변경이력 :
	 * 
	 * @param InputStream is, POIFiletypes POIFileTypes, Class<T> classType
	 * @return List<T>
	 */
	public static <T> List<T> parseRegularExcel(InputStream is, POIFiletypes POIFileTypes, Class<T> classType){		
		List<T> resultList = null;
		try {
//			getWorkbookObj(is, type);
			resultList = (List<T>) getObjFromParsingExcel(is, POIFileTypes, classType);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return resultList;
	}
		
	private static <T> List<T> getObjFromParsingExcel(InputStream is, POIFiletypes type, Class<T> classType) throws Exception{
		List<T> beanList = new ArrayList<T>();
		List<String> propertyNames = new ArrayList<String>();
		Field[] fields = classType.getDeclaredFields();
		
		if(type.equals(POIFiletypes.XLS)){
			POIFSFileSystem fs = null;
			try{
				fs = new POIFSFileSystem(is);
			}catch(Exception e){
				throw new IllegalArgumentException("fail convert inputstream type to POIFSFileSystem type");
			}
			HSSFWorkbook hssfWorkbook = new HSSFWorkbook(fs);
			int sheetNum =  hssfWorkbook.getNumberOfSheets();
			
			SheetLoop : for(int i=0; i<sheetNum; i++){
				HSSFSheet sheet = hssfWorkbook.getSheetAt(i);
				int rows = sheet.getLastRowNum();
				for(int k=0; k<rows; k++){
					HSSFRow row = sheet.getRow(k);
					
					// HSSFSheet 클래스에서는 excel sheet의 끝을 알려주는(data가 없는 row) 메소드가 없으므로 이와같이 row의 끝을 체크합니다.
					if(row == null){
						continue SheetLoop;
					}
					
					int cells = row.getPhysicalNumberOfCells();
					
					if(k==0){
						for(int c=0; c<cells; c++){
							HSSFCell cell = row.getCell(c);
							if(cell == null)
								continue;
							propertyNames.add(cell.getStringCellValue());
						}
						continue;
					}
					T dataBindingObj = classType.getConstructor().newInstance();
					for(int c=0; c<cells; c++){
						HSSFCell cell = row.getCell(c);
						if(cell == null)
							continue;
						switch(cell.getCellType()){
							case Cell.CELL_TYPE_NUMERIC:
								if(DateUtil.isCellDateFormatted(cell)){
									PropertyUtils.setProperty(dataBindingObj, propertyNames.get(c), new Timestamp(cell.getDateCellValue().getTime()));
								}else{
									PropertyUtils.setProperty(dataBindingObj, propertyNames.get(c), cell.getNumericCellValue());
								}
								break;
							case Cell.CELL_TYPE_STRING:
								PropertyUtils.setProperty(dataBindingObj, propertyNames.get(c), cell.getStringCellValue());
								break;
							case Cell.CELL_TYPE_FORMULA :
								PropertyUtils.setProperty(dataBindingObj, propertyNames.get(c), cell.getCellFormula());
								break;
							case Cell.CELL_TYPE_BLANK :
								PropertyUtils.setProperty(dataBindingObj, propertyNames.get(c), null);
								break;
							case Cell.CELL_TYPE_BOOLEAN :
								PropertyUtils.setProperty(dataBindingObj, propertyNames.get(c), cell.getBooleanCellValue());
								break;
						}
					}
					beanList.add(dataBindingObj);
				}
			}
		}else if(type.equals(POIFiletypes.XLSX)){
			XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
			int sheetNum = xssfWorkbook.getNumberOfSheets();
			SheetLoop : for(int i=0; i<sheetNum; i++){
				XSSFSheet sheet = xssfWorkbook.getSheetAt(i);
				
				//int rows = sheet.findEndOfRowOutlineGroup(0);
				//int rows = sheet.getLastRowNum();
				
				// 간혹 특정 excel 파일로부터 row의 갯수를 못읽어 오거나 마지막 row의 위치를 잘못 읽어오는 경우가 있음.
				// 따라서 위처럼 메소드를 호출하여 row의 갯수를 불러오기보다는 아래와 같이 row가 비어있는 지를 체크하도록 함.
				for(int k=0; ; k++){
					XSSFRow row = sheet.getRow(k);
					
					if(row == null)
						continue SheetLoop;
					
					int cells = row.getPhysicalNumberOfCells();
					
					if(k==0){
						for(int c=0; c<cells; c++){
							XSSFCell cell = row.getCell(c);
							if(cell == null){
								continue;
							}
							propertyNames.add(cell.getStringCellValue());
						}
						continue;
					}
					T dataBindingObj = classType.getConstructor().newInstance();
					for(int c=0; c<cells; c++){	
						for(Field field : fields){
							XSSFCell cell = row.getCell(c);
							if(cell == null){
								continue;
							}
							if(field.getName().equals(propertyNames.get(c))){
								switch(cell.getCellType()){								
									case Cell.CELL_TYPE_NUMERIC:
										if(DateUtil.isCellDateFormatted(cell)){
											PropertyUtils.setProperty(dataBindingObj, propertyNames.get(c), new Timestamp(cell.getDateCellValue().getTime()));
										}else if(PropertyUtils.getPropertyType(dataBindingObj, propertyNames.get(c)) == int.class){
											PropertyUtils.setProperty(dataBindingObj, propertyNames.get(c), (int)cell.getNumericCellValue());	// return int
										}else if(PropertyUtils.getPropertyType(dataBindingObj, propertyNames.get(c)) == short.class){
											PropertyUtils.setProperty(dataBindingObj, propertyNames.get(c), (short)cell.getNumericCellValue());	// return short
										}else if(PropertyUtils.getPropertyType(dataBindingObj, propertyNames.get(c)) == long.class){
											PropertyUtils.setProperty(dataBindingObj, propertyNames.get(c), (long)cell.getNumericCellValue());	// return long
										}else if(PropertyUtils.getPropertyType(dataBindingObj, propertyNames.get(c)) == BigDecimal.class){
											PropertyUtils.setProperty(dataBindingObj, propertyNames.get(c), BigDecimal.valueOf(cell.getNumericCellValue()));	// return BigDecimal
										}else{
											PropertyUtils.setProperty(dataBindingObj, propertyNames.get(c), cell.getNumericCellValue());	// return double
										}
										break;
									case Cell.CELL_TYPE_STRING : 
										PropertyUtils.setProperty(dataBindingObj, propertyNames.get(c), cell.getStringCellValue());		// return string
										break;
									case Cell.CELL_TYPE_FORMULA :
										PropertyUtils.setProperty(dataBindingObj, propertyNames.get(c), cell.getCellFormula());			// return string
										break;
									case Cell.CELL_TYPE_BLANK :
										//System.out.println("cell is null");
										break;
									case Cell.CELL_TYPE_BOOLEAN :
										PropertyUtils.setProperty(dataBindingObj, propertyNames.get(c), cell.getBooleanCellValue());	// return boolean
										break;
								}
							}
						}
					}
					beanList.add(dataBindingObj);
				}
			}
		}else{
			throw new Exception("Excel mapping is no there");
		}
		return beanList;
	}
	
	public static void main(String[] args) throws Throwable {
		InputStream is = new FileInputStream(new File("C:/test_excel.xlsx"));
		List<TestExcel> testExcelObjs = ExcelParserUsePOI.getObjFromParsingExcel(is, POIFiletypes.XLSX, TestExcel.class);
		for (TestExcel testExcel : testExcelObjs) {
			System.out.println(testExcel);
		}
	}
}