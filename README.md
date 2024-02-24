# commonPoiUtils
Excel模板导出工具

通过使用占位符，快速导出数据列表

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

    ![TU(T9JVF_9}){1B4Z``MGUH](https://github.com/996lsz/commonPoiUtils/assets/49548423/a992efc9-4287-4dc4-b87b-6afea92d2440)
![图片](https://github.com/996lsz/commonPoiUtils/assets/49548423/f183cf62-3875-45d3-8591-7e7d3e2596bd)
