package cn.lsz.commonpoiutils;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

class CommonPoiUtilsApplicationTests {

	public static void main(String[] args) throws Exception {
		Resource resource = new ClassPathResource("导出模板.xlsx");

		List<Employee> employeeList = new ArrayList<>();

		Dept dept1 = new Dept("1", "技术部");
		Dept dept2 = new Dept("2", "产品部");

		employeeList.add(new Employee("1", "张三", 25, dept1));
		employeeList.add(new Employee("2", "李四", 26, dept1));
		employeeList.add(new Employee("1", "王五", 27, dept2));

		try(
				XSSFWorkbook commonExcel = CommonPoiUtils.createCommonExcel(resource.getInputStream(), employeeList);
				FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\lsz\\Desktop\\导出结果.xlsx");

		){
			commonExcel.write(fileOutputStream);
		}


	}


}
