package cn.lsz.commonpoiutils;

public class Employee {

    private String id;

    private String name;

    private Integer age;

    private Dept dept;

    public Employee(){}

    public Dept getDept() {
        return dept;
    }

    public void setDept(Dept dept) {
        this.dept = dept;
    }

    public Employee(String id, String name, Integer age, Dept dept){
        this.id = id;
        this.name = name;
        this.age = age;
        this.dept = dept;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }
}
