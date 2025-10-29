package com.example.excel;

public class excelFileClass {
    String name;
    String path;
    sheetClass[] sheets;

    public String getName() {
        return name;
    }

    public String getPath() {
        return path;
    }

    public sheetClass[] getSheets() {
        return sheets;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setPath(String path) {
        this.path = path;
    }

    public void setSheets(sheetClass[] sheets) {
        this.sheets = sheets;
    }

    public excelFileClass(String name, String path) {
        this.name = name;
        this.path = path;

        
    } 
}
