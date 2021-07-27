package com.example.export_excel_using_template;

import java.math.BigDecimal;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExportExcelUtil {

    public static void createSheet(HSSFWorkbook workbook, Integer sheetNum, String sheetTitle, String title,
            String[] rowName, List<List<String>> dataList) throws Exception {
        try {
            // Create excel sheet
            HSSFSheet sheet = workbook.createSheet(sheetTitle);

            // Table header style
            HSSFCellStyle columnTopStyle = getColumnTopStyle(workbook);

            // Body style
            HSSFCellStyle style = getStyle2(workbook);

            // Number of column
            int columnNum = rowName.length;

            // Create first row data list(header)
            HSSFRow rowRowName = sheet.createRow(0);
            for(int n=0;n<columnNum;n++){
                // Create cell column
                HSSFCell  cellRowName = rowRowName.createCell(n);
                // Set column type
                cellRowName.setCellType(HSSFCell.CELL_TYPE_STRING);
                // Convert data to excel data type
                HSSFRichTextString text = new HSSFRichTextString(rowName[n]);
                // Put data to column
                cellRowName.setCellValue(text);
                // Set column style
                cellRowName.setCellStyle(columnTopStyle);
            }

            // Create data list(body)
            for(int i=0;i<dataList.size();i++){
                List<String> obj = dataList.get(i);
                HSSFRow row = sheet.createRow(i+1);
                for(int j=0; j<obj.size(); j++){
                    HSSFCell  cell = null;
                    // create cell at which place and set the type
                    if (j == 0) {
                        cell = row.createCell(j,HSSFCell.CELL_TYPE_STRING);
                    } else {
                        cell = row.createCell(j,HSSFCell.CELL_TYPE_NUMERIC);
                    }
                    // data checking
                    if(!"".equals(obj.get(j)) && obj.get(j) != null){
                        if (j == 0) {
                            cell.setCellValue(obj.get(j).toString());
                        } else {
                            cell.setCellValue(Double.parseDouble(obj.get(j).toString()));
                        }
                    }else {
                        cell.setCellValue(" ");
                    }
                    cell.setCellStyle(style);
                }
            }

            // Set the column width according to the length of cell value
            for (int colNum = 0; colNum < columnNum; colNum++) {
                int columnWidth = sheet.getColumnWidth(colNum) / 256;
                for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
                    HSSFRow currentRow;
                    // havent been used
                    if (sheet.getRow(rowNum) == null) {
                        currentRow = sheet.createRow(rowNum);
                    } else {
                        currentRow = sheet.getRow(rowNum);
                    }
                    if (currentRow.getCell(colNum) != null) {
                        HSSFCell currentCell = currentRow.getCell(colNum);
                        if (currentCell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
                            if (currentCell.getRichStringCellValue() == null) {
                                continue;
                            }
                            int length = currentCell.getStringCellValue().getBytes().length;
                            if (columnWidth < length) {
                                columnWidth = length;
                            }
                        }
                    }
                }
                sheet.setColumnWidth(colNum, 18 * 256);
            }
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    public static HSSFCellStyle getColumnTopStyle(HSSFWorkbook workbook) {
        // create font
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 11);
        font.setFontName("Arial");
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        // style.setWrapText(true);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        return style;
    }

    public static HSSFCellStyle getStyle2(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setFontName("Arial");
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        // style.setWrapText(true);
        // style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        return style;
    }
}
