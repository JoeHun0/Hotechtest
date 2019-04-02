package Primary;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.*;

public class Tasks {
    public void InputFromExcell(File be) {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(be));
            Input input = new Input();
            input.databaseTorol();
            input.insertThom(workbook);
            input.insertRhom(workbook);
            input.insertTuasz1(workbook);
            input.insertTuag1(workbook);
            input.insertTuasz2(workbook);
            input.insertTuag2(workbook);
            input.insertHatfok(workbook);
            input.insertGozpar(workbook);
            input.insertVizpar(workbook);
            input.insertTuztadat(workbook);
            input.insertVegyes(workbook);
            input.insertLogika(workbook);
        } catch (FileNotFoundException e) {
            System.out.println(e.getMessage());
        } catch (IOException fe) {
            System.out.println(fe.getMessage());
        }
        System.out.println("finish");
    }

}
