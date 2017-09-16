package com.excelimport;

import com.excelimport.model.People;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import javax.servlet.http.HttpServletResponse;
/**
 * Created by oguz on 14.09.2017.
 */
@RestController
@RequestMapping
public class ExportController {

    @GetMapping(value = "export")
    @ResponseBody
    public void export () {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("People");
        Object[][] datatypes = {
                {"Datatype", "Type", "Size(in bytes)"},
                {"int", "Primitive", 2},
                {"float", "Primitive", 4},
                {"double", "Primitive", 8},
                {"char", "Primitive", 1},
                {"String", "Non-Primitive", "No fixed size"}
        };
        People p1 = new People()
                .setLastName("yilmaz")
                .setName("oguz");
        People p2 = new People()
                .setLastName("erik")
                .setName("yağmur");
        List<People> peopleList = new ArrayList<>();
        peopleList.add(p1);
        peopleList.add(p2);
        int rowNum = 0;
        System.out.println("Creating excel");
        XSSFCellStyle style = workbook.createCellStyle();
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillBackgroundColor(IndexedColors.RED.getIndex());

        for (People people : peopleList) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            for (int i = 0; i<1; i++) {
                Cell cell = row.createCell(colNum++);
                cell.setCellStyle(style);
                cell.setCellValue(people.getName());
                cell = row.createCell(colNum++);
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setFillBackgroundColor(IndexedColors.BLUE.getIndex());
                cell.setCellStyle(style);
                cell.setCellValue(people.getLastName());
            }
        }

        try {
            FileOutputStream outputStream = new FileOutputStream("people.xlsx");
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");
    }
    @GetMapping(value = "exportExcel")
    @ResponseBody
    public void downloadXLS(HttpServletResponse response) {

        try {
            response.setContentType("application/vnd.ms-excel");
            String reportName = "test.xls";
            response.setHeader("Content-disposition", "attachment; filename=" + reportName);
            ArrayList<String> rows = new ArrayList<String>();
            rows.add("Is Yeri");
            rows.add("\t");
            rows.add("Lokasyon");
            rows.add("\t");
            rows.add("Faliyetin Tanimi");
            rows.add("\t");
            rows.add("Tehlike Kaynagi");
            rows.add("\t");
            rows.add("Olasi Risk");
            rows.add("\t");
            rows.add("Tehlikeli Durum Ve Davranislar");
            rows.add("\t");
            rows.add("Risk Degerlendirme Tarihi");
            rows.add("\t");
            rows.add("Maruziyet");
            rows.add("\t");
            rows.add("Olasilik");
            rows.add("\t");
            rows.add("Frekans");
            rows.add("\t");
            rows.add("Siddet");
            rows.add("\t");
            rows.add("Risk Skoru");
            rows.add("\t");
            rows.add("Risk Degeri");
            rows.add("\t");
            rows.add("Planlanan Faaliyetler");
            rows.add("\t");
            rows.add("Planlanan Faaliyetin Sorumlusu");
            rows.add("\t");
            rows.add("Planlanan Faaliyetin Gerceklesme Suresi");
            rows.add("\t");
            rows.add("DOF Sonrasi Olasilik");
            rows.add("\t");
            rows.add("DOF Sonrasi Frekans");
            rows.add("\t");
            rows.add("DOF Sonrasi Siddet");
            rows.add("\t");
            rows.add("DOF Sonrasi Risk Skoru");
            rows.add("\t");
            rows.add("Fotograf Sec");
            rows.add("\n");

            for (int i = 0; i < 1; i++) {
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\t");
                rows.add("data");
                rows.add("\n");
            }
            Iterator<String> iter = rows.iterator();
            while (iter.hasNext()) {
                String outputString = (String) iter.next();
                response.getOutputStream().print(outputString);
            }

            response.getOutputStream().flush();

        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
