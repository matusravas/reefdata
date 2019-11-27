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

    StringBuilder fileName;
    File folder;
    String fileWithPath;

    void findExcelFiles(String name) {
        fileName = new StringBuilder(name);
        folder = new File("./");
        if (name.isEmpty()) {
            System.out.println("Zoznam najdenych suborov: ");
            File[] listOfFiles = folder.listFiles();
            for (int i = 0; i < listOfFiles.length; i++) {
                if (listOfFiles[i].getName().contains(".xlsx")) {
                    if (listOfFiles[i].isFile() && listOfFiles[i].getName().contains("body")) {
                        File file = new File(listOfFiles[i].getAbsolutePath());
                        System.out.println(i + ": " + file);
                        try {
                            fileIn = new FileInputStream(file);
                        } catch (FileNotFoundException e) {
                            e.printStackTrace();
                        }
                    }
                }
            }
            BufferedReader buffer = new BufferedReader(new InputStreamReader(System.in));
            System.out.print("Zadajte poradove cislo z predchadzajuceho zoznamu suborov: ");
            try {
                int fileIndex = Integer.parseInt(buffer.readLine());
                fileName.append(listOfFiles[fileIndex]);
                fileWithPath = folder.getAbsolutePath() + "\\" + fileName;
                fileIn = new FileInputStream(fileWithPath);
                System.out.println(listOfFiles[fileIndex]);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            folder = new File("");
            try {
                if (!name.contains(".")) {//&& !name.substring(".xlsx".length() - 5).equals(".xlsx")) {
                    fileName.append(".xlsx");
                }
                fileWithPath = folder.getAbsolutePath() + "\\" + fileName;
                System.out.println(fileWithPath);
                fileIn = new FileInputStream(new File(fileWithPath));
//                fileIn = new FileInputStream(new File("C:\\BOX\\ReefData\\bodyx.xlsx")); //alternativa natvrdo
            } catch (FileNotFoundException e) {
                System.out.println(folder.getAbsolutePath() + "\\" + fileName + " neexistuje");
                e.printStackTrace();
            }
            System.out.println(fileName);
        }
    }

    void openExcelDoc() {
        System.out.println("Opening document: " + fileWithPath);
        try {
//            fileIn = new FileInputStream(new File("C:\\BOX\\ReefData\\bodyx.xlsx"));
            //fileIn = new FileInputStream(new File("C:\\BOX\\ReefData\\out\\artifacts\\ReefData_jar\\bodyx.xlsx"));
            workbook = new XSSFWorkbook(fileIn);
        } catch (IOException e) {
            System.out.println("Dokument bodyx.xlsx neexistuje v adresari, kde sa nachadza spustany .exe subor!");
        }
        try {
            gpsSheet = workbook.getSheetAt(0);
            sonarSheet = workbook.getSheetAt(1);
            timeDelaySheet = workbook.getSheetAt(2);
        } catch (IllegalArgumentException ex) {
            System.out.println("Skontroluj spravnost harkov!\n-Prvy harok GPS body\n-Druhy harok sonar body\n-Treti harok casovy posun");

        }
    }

    void closeExcelDoc() {
        try {
            fileIn.close();
        } catch (IOException e) {
            System.out.println("Dokument body.xlsx je prave otovreny, zatvor ho a spusti .exe znova!");
//            e.printStackTrace();
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
                System.out.println("V 1. harku (GPS) musia byt GPS suradnice x-ove v 2.stlpci, suradnice y-ove v 3. stlpci a casy v 12. stlpci!!!");

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
//                    cal.add(Calendar.HOUR_OF_DAY, this.timeDelay);
                    cal.add(Calendar.HOUR_OF_DAY, -(int) timeDelayHours);
                    cal.add(Calendar.MINUTE, -(int) timeDelayMinutes);
                    cal.add(Calendar.SECOND, -(int) timeDelaySeconds);
                    data.setTimestamp(cal.getTimeInMillis());
                } catch (ParseException e) {
                    e.printStackTrace();
                }
                sonarDataList.add(data);
            } catch (IllegalArgumentException ex) {
                System.out.println("V 2. harku (sonar) musia byt hlbky v 3.stlpci a casy v 8. stlpci!!!");
            }
        }
    }

    /**
     * Stlpec za GPS data bude rozdiel GPS cas - sonar cas
     */
    void writeValuesToExcel() {
        long currentGPS;
        long currentSonar;
        long nextSonar;
        int sonarIndex = 0;
        Calendar cal = Calendar.getInstance();
        SimpleDateFormat df = new SimpleDateFormat("HH:mm:ss.SS");
        try {
            fileOut = new FileOutputStream(new File(folder.getAbsolutePath() + "\\" + fileName));
//            fileOut = new FileOutputStream(new File("body.xlsx"));
            removeDataSheet(workbook);
            dataSheet = workbook.createSheet("data");
        } catch (FileNotFoundException e) {
            System.out.println("Hodnoty sa nedaju zapisat, pretoze dokument body.xlsx je prave otovreny, zatvor ho a spusti .exe znova!");
//            e.printStackTrace();
        }
        for (int i = 0; i < gpsDataList.size(); i++) { //prechadzam data z GPS
            currentGPS = gpsDataList.get(i).getTimestamp();
            for (int j = 0; j < sonarDataList.size(); j++) { //hladam data zo sonarru pre najblizsi cas sonaru ku GPS
                currentSonar = sonarDataList.get(j).getTimestamp();
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
//            rowIndex++;
        }
        try {
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
        } catch (IOException e) {
            System.out.println("Hodnoty sa nedaju zapisat, pretoze dokument body.xlsx je prave otovreny, zatvor ho a spusti .exe znova!");
        }
    }

    void readTimeDelay() {
        try {
            Row row = timeDelaySheet.getRow(1);
            this.timeDelayHours = row.getCell(0).getNumericCellValue();
            this.timeDelayMinutes = row.getCell(1).getNumericCellValue();
            this.timeDelaySeconds = row.getCell(2).getNumericCellValue();
        } catch (IllegalArgumentException ex) {
            System.out.println("V 3. harku (casovy posun) musi byt v 1. riadku a 1.stlpec celociselna hodnota casoveho posunu!!!");

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
}


