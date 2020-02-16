package com.mrdev;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

class Excel {
    private FileInputStream fileIn;
    private FileOutputStream fileOut;
    private XSSFWorkbook workbook;
    private XSSFSheet gpsSheet;
    private XSSFSheet sonarSheet;
    private XSSFSheet dataSheet;
    private XSSFSheet timeDelaySheet;


    private double timeDelayHours;
    private double timeDelayMinutes;
    private double timeDelaySeconds;

    private ArrayList<GPSData> gpsDataList;
    private List<SonarData> sonarDataList;

    private StringBuilder fileName;
    private String fileWithPath;

    /**
     * @param name filename
     * @return 0 if file exist 1 if not
     */
    int findExcelFiles(String name) {
        File folder;
        if (name.isEmpty()) {
            folder = new File("./");
            System.out.println("Zoznam najdenych suborov:\n");
            File[] listOfFiles = folder.listFiles();
            ArrayList<File> listOfWantedFiles = new ArrayList<>();
            int countOfWantedFiles = 0;
            if (listOfFiles != null) {
                for (File listOfFile : listOfFiles) {
                    if (listOfFile.isFile() && listOfFile.getName().contains(".xlsx")) {
                        File file = new File(listOfFile.getName());
                        listOfWantedFiles.add(file);
                        countOfWantedFiles++;
                        System.out.println(countOfWantedFiles + ": " + listOfWantedFiles.get(countOfWantedFiles - 1));
                        try {
                            fileIn = new FileInputStream(file);
                        } catch (FileNotFoundException e) {
                            e.printStackTrace();
                        }
                    }
                }
            }
            BufferedReader buffer = new BufferedReader(new InputStreamReader(System.in));
            System.out.print("\nZadajte poradove cislo z predchadzajuceho zoznamu suborov: ");
            try {
                fileName = new StringBuilder();
                int index = Integer.parseInt(buffer.readLine());
                fileName.append(listOfWantedFiles.get(index - 1).getName());
                fileWithPath = folder.getAbsolutePath() + "\\" + fileName;
                fileIn = new FileInputStream(fileWithPath);
            } catch (IOException e) {
                return 1;
            }
        } else {
            fileName = new StringBuilder(name);
            folder = new File("");
            try {
                if (!name.contains(".")) {
                    fileName.append(".xlsx");
                } else if (name.contains(".")) {
                    int indexOfDot = fileName.indexOf(".");
                    String n = fileName.substring(0, indexOfDot);
                    fileName.replace(0, fileName.length(), n + ".xlsx");
                }
                fileWithPath = folder.getAbsolutePath() + "\\" + fileName;
                fileIn = new FileInputStream(new File(fileWithPath));
            } catch (FileNotFoundException e) {
                System.out.println(folder.getAbsolutePath() + "\\" + fileName + " neexistuje");
                return 1;
            }
        }
        return 0;
    }

    void openExcelDoc() {
        System.out.println("Opening document: " + fileWithPath);
        try {
            workbook = new XSSFWorkbook(fileIn);
        } catch (IOException e) {
            System.out.println("Dokument " + fileWithPath + " neexistuje v adresari, kde sa nachadza spustany .exe subor!");
        }
        try {
            this.findRelevantSheets();
        } catch (IllegalArgumentException ex) {
            System.out.println("Skontroluj ci dostupnost obsahuje vsetky potrebne harky:\n-harok: gps \n-harok: sonar body\n-harok: posun");

        }
    }

    void closeExcelDoc() {
        try {
            fileIn.close();
        } catch (IOException e) {
            System.out.println("Dokument " + fileWithPath + " je prave otovreny, zatvor ho a spusti .exe znova!");
        }
    }

    void readGPSData() {
        gpsDataList = new ArrayList<>();
        DataFormatter dataFormatter = new DataFormatter();
        SimpleDateFormat df = new SimpleDateFormat("HH:mm:ss.SS");
        for (int i = 0; i < gpsSheet.getLastRowNum() + 1; i++) {
            try {
                Row row = gpsSheet.getRow(i);
                if (row.getCell(0) != null) {
                    Cell name = row.getCell(0);
                    Cell lat = row.getCell(1);
                    Cell lon = row.getCell(2);
                    GPSData data = new GPSData();
                    data.setName(name.getStringCellValue());
                    data.setLatitude(lat.getNumericCellValue());
                    data.setLongitude(lon.getNumericCellValue());
                    try {
                        data.setTimestamp(df.parse(dataFormatter.formatCellValue(row.getCell(13))).getTime());
                    } catch (ParseException e) {
                        e.printStackTrace();
                    }
                    gpsDataList.add(data);
                }
            } catch (IllegalArgumentException ex) {
                System.out.println("V harku (GPS) musia byt x-ove suradnice v 2.stlpci, y-ove v 3. stlpci a casy v 12. stlpci!!!");

            }
        }
    }

    void readSonarData() {
        this.readTimeDelay();
        sonarDataList = new ArrayList<>();
        Calendar cal = Calendar.getInstance();
        DataFormatter dataFormatter = new DataFormatter();
        SimpleDateFormat df = new SimpleDateFormat("HH:mm:ss");
        for (int i = 0; i < sonarSheet.getPhysicalNumberOfRows(); i++) {
            try {
                Row row = sonarSheet.getRow(i);
                Cell depth = row.getCell(2);
                SonarData data = new SonarData();
                data.setDepth(depth.getNumericCellValue());
                try {
                    cal.setTime(df.parse(dataFormatter.formatCellValue(row.getCell(7))));
                    cal.add(Calendar.HOUR_OF_DAY, -(int) timeDelayHours);
                    cal.add(Calendar.MINUTE, -(int) timeDelayMinutes);
                    cal.add(Calendar.SECOND, -(int) timeDelaySeconds);
                    data.setTimestamp(cal.getTimeInMillis());
                } catch (ParseException e) {
                    e.printStackTrace();
                }
                sonarDataList.add(data);
            } catch (IllegalArgumentException ex) {
                System.out.println("V harku (sonar) musia byt hlbky v 3.stlpci a casy v 8. stlpci!!!");
            }
        }
    }

    /**
     * Stlpec za GPS data bude rozdiel GPS cas - sonar cas
     */
    void writeValuesToExcel() {
        System.out.println("Writing data to file: " + fileName);
        long currentGPS;
        long currentSonar;
        long nextSonar;
        int sonarIndex = 0;
        Calendar cal = Calendar.getInstance();
        SimpleDateFormat df = new SimpleDateFormat("HH:mm:ss.SS");
        try {
            fileOut = new FileOutputStream(new File(fileWithPath));
            removeDataSheet(workbook);
            dataSheet = workbook.createSheet("data");
        } catch (FileNotFoundException e) {
            System.out.println("Hodnoty sa nedaju zapisat, pretoze dokument " + fileWithPath + " je prave otovreny, zatvor ho a spusti .exe znova!");
        }
        for (int i = 0; i < gpsDataList.size(); i++) { //prechadzam data z GPS
            currentGPS = gpsDataList.get(i).getTimestamp();
            for (int j = 0; j < sonarDataList.size(); j++) { //hladam data zo sonarru pre najblizsi cas sonaru ku GPS
                currentSonar = sonarDataList.get(j).getTimestamp();
                System.out.println("i: " + i + " j:" + j);
                if (j != sonarDataList.size() - 1) {
                    nextSonar = sonarDataList.get(j + 1).getTimestamp();
                    if (Math.abs(currentGPS - currentSonar) < Math.abs(currentGPS - nextSonar)) {
                        sonarIndex = j;
                        break;
                    } else sonarIndex = j + 1;
                }
            }
            //writing data excel doc
            Row row = dataSheet.createRow(i);
            Cell title = row.createCell(0);
            title.setCellValue(gpsDataList.get(i).getName());
            Cell lat = row.createCell(1);
            lat.setCellValue(gpsDataList.get(i).getLatitude());
            Cell lon = row.createCell(2);
            lon.setCellValue(gpsDataList.get(i).getLongitude());
            Cell depth = row.createCell(3);
            depth.setCellValue(sonarDataList.get(sonarIndex).getDepth());
            Cell time = row.createCell(4);
            cal.setTimeInMillis(gpsDataList.get(i).getTimestamp());
            time.setCellValue(df.format(cal.getTime()));
            Cell timeDiff = row.createCell(5);
            timeDiff.setCellValue((gpsDataList.get(i).getTimestamp() -
                    sonarDataList.get(sonarIndex).getTimestamp()) / 1000.000);
        }
        try {
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
        } catch (IOException e) {
            System.out.println("Hodnoty sa nedaju zapisat, pretoze dokument " + fileWithPath + "je prave otovreny, zatvor ho a spusti .exe znova!");
        }
    }

    void readTimeDelay() {
        try {
            Row row = timeDelaySheet.getRow(1);
            this.timeDelayHours = row.getCell(0).getNumericCellValue();
            this.timeDelayMinutes = row.getCell(1).getNumericCellValue();
            this.timeDelaySeconds = row.getCell(2).getNumericCellValue();
        } catch (IllegalArgumentException | NullPointerException ex) {
            System.out.println("V harku s casovym posunom (harok posun), musia byt v 2. riadku stlpce:\n");
            System.out.println("1. Hodiny, 2. Minuty, 3. Sekundy");

        }
    }

    private void removeDataSheet(XSSFWorkbook workbook) {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            XSSFSheet tmpSheet = workbook.getSheetAt(i);
            if (tmpSheet.getSheetName().equals("data")) {
                workbook.removeSheetAt(i);
            }
        }
    }

    private void findRelevantSheets() {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            XSSFSheet tmpSheet = workbook.getSheetAt(i);
            switch (tmpSheet.getSheetName().toLowerCase()) {
                case "gps":
                    gpsSheet = workbook.getSheetAt(i);
                    break;
                case "sonar":
                    sonarSheet = workbook.getSheetAt(i);
                    break;
                case "posun":
                    timeDelaySheet = workbook.getSheetAt(i);
                    break;
            }
        }
    }
}


