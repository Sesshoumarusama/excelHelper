package excelHelper.readTest;

import com.pengpei.excelHelper.reader.ExcelReader;
import com.pengpei.excelHelper.reader.FileType;
import com.pengpei.excelHelper.reader.ReadModel;
import com.pengpei.excelHelper.reader.Reader;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * Created by pengpei on 2017/8/29.
 */
public class ReadTestCase {
    @Test
    public void test1() throws IOException {
        InputStream is = new FileInputStream("E:\\file\\students.xlsx");
        Workbook workbook = new XSSFWorkbook(is);
        System.out.println("sheet的数量：" + workbook.getNumberOfSheets());
        Sheet sheet = workbook.getSheetAt(0);
        System.out.println(sheet.getPhysicalNumberOfRows());
        System.out.println(sheet.getLastRowNum());
        Row row = sheet.getRow(0);
        System.out.println(row.getLastCellNum());//从开始1开始
    }

    @Test
    public void test2() throws IOException, IllegalAccessException, InstantiationException {
        InputStream is = new FileInputStream("E:\\file\\students.xlsx");//模拟一个输入流，该输入流可以来源于网络，例如使用SpringMVC上传的文件
        Reader excelReader = ExcelReader.ReaderBuilder.newBuilder().setReadModel(ReadModel.TopToBottom).setFileType(FileType.XLSX).build();
        List<Student> read = excelReader.read(is, Student.class);
        for (Student s : read){
            System.out.println(s);
        }
    }

    @Test
    public void test3() throws IOException, IllegalAccessException, InstantiationException {
        Reader excelReader = ExcelReader.ReaderBuilder.newBuilder().setFileType(FileType.XLSX).build();
        List<Student> read = excelReader.read("E:\\file\\students1.xlsx", Student.class);
        for (Student s : read){
            System.out.println(s);
        }
    }


}
