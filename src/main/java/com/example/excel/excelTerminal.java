package com.example.excel;

import java.util.Scanner;

public class excelTerminal {
    excelFileClass file = new excelFileClass("excel", "");
    public static void main(String[] args) {
        StarterDialog();

        
    }

    private static void StarterDialog() {
        Scanner scn = new Scanner(System.in);

        System.out.println("Wellcome to Exccel File Configurator");
        System.out.print("Enter how many sheets you want in your Excel file: ");
        final int sheetNum = scn.nextInt();

        if(sheetNum == 1) {
            System.out.print("Please incert ");
        }

    }
}
