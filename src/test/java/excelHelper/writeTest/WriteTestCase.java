package excelHelper.writeTest;

import com.pengpei.excelHelper.reader.FileType;
import com.pengpei.excelHelper.reader.ReadModel;
import com.pengpei.excelHelper.writer.ExcelWriter;
import com.pengpei.excelHelper.writer.Writer;
import excelHelper.readTest.Student;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Created by pengpei on 2017/9/2.
 */
public class WriteTestCase {
    @Test
    public void test1() throws Exception {
        Writer writer = ExcelWriter.WriterBuilder.newWriter().
                setFileType(FileType.XLSX).setModel(ReadModel.TopToBottom).build();
        List<Student> students = new ArrayList<>();
        students.add(new Student(1, "张三", 21, new Date(), false));
        students.add(new Student(2, "李四", 22, new Date(), true));
        students.add(new Student(3, "王五", 23, new Date(), false));
        students.add(new Student(4, "赵六", 24, new Date(), true));
        students.add(new Student(5, "钱七", 25, new Date(), false));

        writer.write("E:\\file\\students" + System.currentTimeMillis() + ".xlsx", students);
    }
}
