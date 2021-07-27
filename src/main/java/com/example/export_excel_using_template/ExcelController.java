package com.example.export_excel_using_template;

import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.TemplateExportParams;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
public class ExcelController {
    
    @RequestMapping(value = "download-excel")
    public void downloadExcel(HttpServletResponse response) {
        //use Excel Template
        // String filePath = FilePath.name("sample.xls");
        // TemplateExportParams params = new TemplateExportParams(filePath);

        // For Self desktop to download the excel file and open without corrupted or crashed
        TemplateExportParams params = new TemplateExportParams("ExcelTemplate/sample.xls");
        Map<String, Object> map = new HashMap<>();
        map.put("data", "Single Data");
        List<Map<String, Object>> list = new ArrayList<>();
        Map<String, Object> tempMap;
        for (int i = 0; i < 2; i++) {
            tempMap = new HashMap<>();
            tempMap.put("name", "name " + i);
            tempMap.put("amount", BigDecimal.valueOf(i).setScale(2));
            list.add(tempMap);
        }
        map.put("dataList", list);

        // Write the workbook
        Workbook workbook = ExcelExportUtil.exportExcel(params, map);

        OutputStream output = null;
        response.setContentType("application/force-download");// 应用程序强制下载

        try {
            response.setHeader("Content-Disposition", "attachment;filename=Result_Data.xls");
            output = response.getOutputStream();
            workbook.write(output);
            output.flush();
        } catch (Exception e) {
            System.out.println(e);
        } finally {
            try {
                if (output != null) {
                    output.close();
                }
            } catch (Exception e) {
                System.out.println(e);
            }
        }
    }

    @RequestMapping(value = "create-excel")
    public void selfCreateExcel(HttpServletResponse response) {
        Map<String, Object> map = new HashMap<>();
        map.put("generatedDate", "here");
        List<List<String>> dataList = new ArrayList<List<String>>();
        List<String> rowList = null;
        for (int i = 0; i < 2; i++) {
            rowList = new ArrayList<String>();
            rowList.add("name " + "name " + "name "+ "name "+ "name "+ "name " + i);
            rowList.add(BigDecimal.valueOf(i).setScale(2).toString());
            dataList.add(rowList);
        }
        String fileName = "Result_Data";
        String[] header = new String[]{"Name", "Amount"};
        OutputStream out = null;
        try{
            response.setContentType("application/force-download");
            response.setHeader("Content-Disposition", "attachment;filename="+fileName+".xls");
            out = response.getOutputStream();
            Workbook wb = downloadExcel(fileName, header, dataList);;
            wb.write(out);
            out.flush();
        }catch (Exception e){
            System.out.println(e);
        }finally {
            try{
                if(out!=null){
                    out.close();
                }
            }catch (Exception e){
                System.out.println(e);
            }
        }
    }

    private HSSFWorkbook downloadExcel(String fileName, String[] header, List<List<String>> dataList){
        HSSFWorkbook wb = new HSSFWorkbook();
        try{
            ExportExcelUtil.createSheet(wb,0,fileName,fileName,header,dataList);
        }catch (Exception e){
            System.out.println(e);
        }
        return wb;
    }
}
