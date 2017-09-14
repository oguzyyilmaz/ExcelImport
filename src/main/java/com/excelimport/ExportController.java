package com.excelimport;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import javax.servlet.http.HttpServletResponse;
/**
 * Created by oguz on 14.09.2017.
 */
@RestController
@RequestMapping
public class ExportController {

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

            for (int i = 0; i < 20; i++) {
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
