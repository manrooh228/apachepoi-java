package com.example.excelExamples;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOError;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {
    private File excelFile;
    private Workbook workbook;
    private Sheet sheet;

    public ExcelWriter(File excelFile) {
        this.excelFile = excelFile;
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Sheet 1");
        sheet.setColumnWidth(2, 15*256);
    }

    public static void main(String[] args) throws IOException {
        File excelFile = Paths.get("").resolve("excel.xlsx").toFile();
        ExcelWriter writer = new ExcelWriter(excelFile);
        writer.writeToExcelFile();
        System.out.println("Excel file created");
    }

    private void writeToExcelFile() throws IOException {
        createHeaderRow(0, "Name", "Age", "Date of birth");

        createContentRow(1, "Maksat", 32, LocalDate.of(1993, 4, 12));

        writeWorkbookToExcelFile();
    }

    private void createHeaderRow(int index, String firstColummn, String secondColumn, String thirdColumn) {
        Row headerRow = sheet.createRow(index);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setBorderBottom(BorderStyle.THIN);

        createHeaderCell(headerRow, 0, firstColummn, headerStyle);
        createHeaderCell(headerRow, 1, secondColumn, headerStyle);
        createHeaderCell(headerRow, 2, thirdColumn, headerStyle);
    }


    private void createHeaderCell(Row headerRow, int index, String value, CellStyle style) {
        Cell headerCell = headerRow.createCell(index);
        headerCell.setCellValue(value);
        headerCell.setCellStyle(style);
    }

    private void createContentRow(int index, String name, int age, LocalDate birthDate) {
        Row contentRow = sheet.createRow(index);

        Cell contentCell = contentRow.createCell(0);
        contentCell.setCellValue(name);

        contentCell = contentRow.createCell(1);
        contentCell.setCellValue(age);

        CellStyle dateContentCellStyle = workbook.createCellStyle();
        short dateFormat = workbook.getCreationHelper().createDataFormat().getFormat("mm-dd-yyyy");
        dateContentCellStyle.setDataFormat(dateFormat);

        contentCell = contentRow.createCell(2);
        contentCell.setCellValue(birthDate);
        contentCell.setCellStyle(dateContentCellStyle);
    }
    private void writeWorkbookToExcelFile() throws IOException {
        FileOutputStream outputStream = new FileOutputStream(excelFile);
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
    }
} 
