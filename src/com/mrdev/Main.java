package com.mrdev;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Calendar;

public class Main {

    public static void main(String[] args) {
        Excel excel = new Excel();
        //finding available excel documents
        String fileName = "";
        BufferedReader buffer = new BufferedReader(new InputStreamReader(System.in));
        do {
            System.out.print("Zadaj nazov excel suboru: ");
            try {
                fileName = buffer.readLine();
            } catch (IOException e) {
                e.printStackTrace();
            }
        } while (excel.findExcelFiles(fileName) != 0);
        //opening the excel document
        long beginTime = getBeginTime();
        excel.openExcelDoc();
        //reading data
        System.out.println("Reading data...");
        excel.readGPSData();
        excel.readTimeDelay();
        excel.readSonarData();
        excel.closeExcelDoc();
        //writing data
        excel.writeValuesToExcel();
        long endTime = getFinishTime();
        System.out.println("Successfully finished.");
        System.out.println("Processing time: " + ((endTime - beginTime) / 1000.000) + " sec.");
        System.out.print("Press any key to exit...");
        try {
            System.in.read();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static long getBeginTime() {
        Calendar cal = Calendar.getInstance();
        return cal.getTimeInMillis();
    }

    private static long getFinishTime() {
        Calendar cal = Calendar.getInstance();
        return cal.getTimeInMillis();
    }
}
