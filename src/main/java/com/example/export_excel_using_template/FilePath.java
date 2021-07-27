package com.example.export_excel_using_template;

public class FilePath {

    static public String name(String templateName) {
        FilePath file = new FilePath();
        String path = file.getClass().getClassLoader().getResource("").getPath()+"ExcelTemplate/"+templateName;
        return path;
    }
}
