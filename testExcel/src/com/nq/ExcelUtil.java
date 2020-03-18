package com.nq;

import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelUtil {

	private static int sheetSize = 5000;

	public static <T> void listToExcel(List<T> data,OutputStream out,Map<String, String> fields) throws Exception {
		if(data == null || data.size() == 0) {
			throw new Exception("Excel中无数据");
		}
		HSSFWorkbook workBook = new HSSFWorkbook();
		int sheetNum = data.size()/sheetSize;
		if(data.size()%sheetSize != 0) {
			sheetNum += 1;
		}
		String[] fieldNames = new String[fields.size()];
		String[] chinaNames = new String[fields.size()];
		int count = 0;
		for(Entry<String,String> entry : fields.entrySet()) {
			String fieldName = entry.getKey();
			String chinaName = entry.getValue();
			fieldNames[count] = fieldName;
			chinaNames[count] = chinaName;
			count++;
		}
		for(int i = 0; i < sheetNum; i++) {
			int rowCount = 0;
			HSSFSheet sheet = workBook.createSheet();
			int startIndex = i*sheetSize;
			int endIndex = (i+1)*sheetSize - 1>data.size()?data.size():(i+1)*sheetSize - 1;
			HSSFRow row = sheet.createRow(rowCount);
			//标题行 第一行
			for(int j = 0 ;j < chinaNames.length; j++) {
				HSSFCell cell = row.createCell(j);
				cell.setCellValue(chinaNames[j]);
			}
			rowCount++;
			for(int index = startIndex; index < endIndex; index++) {
				T item = data.get(index);
				row = sheet.createRow(rowCount);
				for(int j = 0; j < chinaNames.length; j++) {
					Field field = item.getClass().getDeclaredField(fieldNames[j]);
					field.setAccessible(true);
					Object o = field.get(item);
					String value = o == null?"":o.toString();
					HSSFCell cell = row.createCell(j);
					cell.setCellValue(value);
				}
				rowCount++;
			}
			workBook.write(out);
		}
		out.close();
		workBook.close();
	}
}
