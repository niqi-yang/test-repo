package com.nq;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class TestExcel {

	public static void main(String[] args) throws Exception {
		List<Student> data = new ArrayList<Student>();
		Student student = new Student();
		student.setId(1);
		student.setAge(20);
		student.setName("二货");
		student.setSex("男");
		student.setHobbies("爱学习");
		data.add(student);student = new Student();
		student.setId(2);
		student.setAge(20);
		student.setName("小李");
		student.setSex("男");
		data.add(student);student = new Student();
		student.setId(3);
		student.setName("小倪");
		data.add(student);student = new Student();
		student.setId(5);
		student.setAge(20);
		student.setName("小王");
		data.add(student);student = new Student();
		student.setId(6);
		student.setAge(20);
		student.setName("张三");
		student.setSex("男");
		data.add(student);student = new Student();
		student.setHobbies("爱学习");
		student.setName("李四");
		student.setSex("男");
		data.add(student);
		OutputStream out = new FileOutputStream("d://student1.xls");
		Map<String, String> fields = new LinkedHashMap<String, String>();
		fields.put("id","编号");
		fields.put("name", "姓名");
		fields.put("age", "年龄");
		fields.put("sex", "性别");
		fields.put("hobbies", "爱好");
		
		ExcelUtil.listToExcel(data, out, fields);
		
		
	}

}
