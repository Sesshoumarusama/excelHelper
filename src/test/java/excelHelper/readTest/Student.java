package excelHelper.readTest;

import com.pengpei.excelHelper.annotation.Cell;
import com.pengpei.excelHelper.annotation.Excel;
import com.pengpei.excelHelper.reader.ReadModel;

import java.util.Date;

/**
 * Created by pengpei on 2017/8/29.
 */
@Excel(model = ReadModel.LeftToRight)
public class Student {
    @Cell(columnNum = "A", rowNum = 1)
    private Integer no;
    @Cell(columnNum = "B", rowNum = 2)
    private String name;
    @Cell(columnNum = "C", rowNum = 3)
    private Integer age;
    @Cell(columnNum = "D", rowNum = 4)
    private Date hiredate;
    @Cell(columnNum = "E", rowNum = 5)
    private Boolean vaild;

    public Student() {
    }

    public Student(Integer no, String name, Integer age, Date hiredate, Boolean vaild) {
        this.no = no;
        this.name = name;
        this.age = age;
        this.hiredate = hiredate;
        this.vaild = vaild;
    }

    @Override
    public String toString() {
        return "Student{" +
                "no=" + no +
                ", name='" + name + '\'' +
                ", age=" + age +
                ", hiredate=" + hiredate +
                ", vaild=" + vaild +
                '}';
    }
}
