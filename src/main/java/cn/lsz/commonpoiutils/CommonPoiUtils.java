package cn.lsz.commonpoiutils;

import org.apache.commons.collections.map.MultiValueMap;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.BeanUtils;

import java.beans.PropertyDescriptor;
import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.Method;
import java.util.Collection;
import java.util.Date;
import java.util.List;
import java.util.regex.Pattern;

/**
 * description
 * 
 * @author LSZ 2020/09/07 9:48
 * @contact 648748030@qq.com
 */
public class CommonPoiUtils {

	private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\{.*\\}");



	public static XSSFWorkbook createCommonExcel(String path, List list) throws Exception {
		try(
			InputStream inputStream = new FileInputStream(path);
		){
			return createCommonExcel(inputStream, list);
		}
	}

	/**
	 * 使用完XSSFWorkbook务必close
	 * @param inputStream 文件
	 * @param list 数据
	 * @return
	 * @throws Exception
	 */
	public static XSSFWorkbook createCommonExcel(InputStream inputStream, List list) throws Exception {
		if(list == null){
			throw new RuntimeException("数据不能为空");
		}
		XSSFWorkbook wb = null;
		try{
			wb = new XSSFWorkbook(inputStream);
			XSSFSheet sheet = wb.getSheetAt(0);
			int i = checkMarked(sheet);
			XSSFRow methodRow = sheet.getRow(i);
			//如果没有数据，删除模板行输出excel结果
			if(list.size() == 0){
				sheet.removeRow(methodRow);
				return wb;
			}
			Class clazz = list.get(0).getClass();
			MultiValueMap rowMethod = getRowMethod(methodRow, clazz);
			//拷贝模板行
			for (int j = 1; j < list.size(); j++) {
				XSSFRow newHeadRow = sheet.createRow(j + i);
				newHeadRow.copyRowFrom(methodRow, new CellCopyPolicy());
			}
			//赋值
			for (int k = 0; k < list.size(); k++) {
				Object object = list.get(k);
				XSSFRow dataRow = sheet.getRow(k + i);
				for (int m = 0; m < rowMethod.entrySet().size(); m++) {
					Collection<Method> methods = rowMethod.getCollection(m);
					if(methods != null) {
						XSSFCell xssfCell = dataRow.getCell(m);
						Object value = object;
						for (Method method : methods) {
							value = method.invoke(value);
						}
						if(value instanceof Date){
							xssfCell.setCellValue(value == null ? "" : DateFormatUtils.format((Date) value, "yyyy-MM-dd HH:mm:ss"));
						}else{
							xssfCell.setCellValue(value == null ? "" : value.toString());
						}
					}
				}
			}
			return wb;
		}catch (Exception e){
			if(wb != null) {
				wb.close();
			}
			if(inputStream != null){
				inputStream.close();
			}
			throw e;
		}

	}

	private static int checkMarked(XSSFSheet sheet) {
		for (int i = 0; i <= sheet.getLastRowNum(); i++) {
			XSSFRow row = sheet.getRow(i);
			int lastCellNum = row.getLastCellNum();
			//获取占位符所在行
			for (int j = 0; j < lastCellNum; j++) {
				XSSFCell cell = row.getCell(j);
				if(cell != null &&  PLACEHOLDER_PATTERN.matcher(cell.getStringCellValue()).matches()){
					return i;
				}
			}
		}
		throw new RuntimeException("没有找到模板方法");
	}

	private static MultiValueMap getRowMethod(XSSFRow lineRow, Class clz) {
		MultiValueMap methods = new MultiValueMap();
		for(int cellNum = 0; cellNum < lineRow.getLastCellNum(); cellNum++) {
			XSSFCell xssfCell = lineRow.getCell(cellNum);
			//只有单元格是方法才加入list
			if (xssfCell != null  && xssfCell.getStringCellValue() != null && PLACEHOLDER_PATTERN.matcher(xssfCell.getStringCellValue()).matches()) {
				String value = xssfCell.getStringCellValue().trim();
				value = value.substring(value.indexOf("{")+1, value.indexOf("}"));
				String[] methodNames = value.split("\\.");
				Class[] clazz = new Class[methodNames.length];
				clazz[0] = clz;
				for (int i = 0; i < methodNames.length; i++) {
					String methodName = methodNames[i];
					PropertyDescriptor propertyDescriptor = BeanUtils.getPropertyDescriptor(clazz[i], methodName);
					Method method = propertyDescriptor.getReadMethod();
					methods.put(cellNum, method);
					if(i < methodNames.length - 1){
						clazz[i + 1] = propertyDescriptor.getPropertyType();
					}
				}

			}else{
				methods.put(cellNum, null);
			}
		}
		return methods;
	}

}
