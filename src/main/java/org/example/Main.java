package org.example;

import io.restassured.response.Response;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import static io.restassured.RestAssured.given;

public class Main {

    public static ThreadLocal<WebDriver> driver = new ThreadLocal<>();
    public static ThreadLocal<Map<String, String>> valueIdMood = ThreadLocal.withInitial(() -> new HashMap<>());
    public static ThreadLocal<Map<String, Integer>> valueMood = ThreadLocal.withInitial(() -> new HashMap<>());

    public static ThreadLocal<Map<Integer, String>> mataUangWeb = ThreadLocal.withInitial(() -> new HashMap<>());

    public static ThreadLocal<Map<Integer, String>> kursJualWeb = ThreadLocal.withInitial(() -> new HashMap<>());

    public static ThreadLocal<Map<Integer, String>> kursBeliWeb = ThreadLocal.withInitial(() -> new HashMap<>());


    public static ThreadLocal<Map<Integer, String>> mataUangAPI = ThreadLocal.withInitial(() -> new HashMap<>());

    public static ThreadLocal<Map<Integer, Double>> kursJualAPI = ThreadLocal.withInitial(() -> new HashMap<>());

    public static ThreadLocal<Map<Integer, Double>> kursBeliAPI = ThreadLocal.withInitial(() -> new HashMap<>());

    public static void main(String[] args) throws IOException, InterruptedException {
        System.setProperty("webdriver.chrome.driver", "C:\\Users\\Lenovo\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");
        driver.set(new ChromeDriver());
        driver.get().manage().window().maximize();
        driver.get().get("https://www.bi.go.id/id/statistik/informasi-kurs/transaksi-bi/default.aspx");
        wait_for_second(2000L);
        getTableValues();
        getAPIKurs();
        compareKurs();
        quit_browser();
//        getAPI();
//        writeToExcel();

    }

    public static void getAPIMood() {
        Response response = given().baseUri("https://680cae692ea307e081d4b97f.mockapi.io").when().get("/mood");

        String responseBody = response.asString();

        if (responseBody.trim().startsWith("[")) {
            JSONArray jsonArray = new JSONArray(responseBody);
            for (int i = 0; i <= 20; i++) {
                JSONObject object = jsonArray.getJSONObject(i);
                valueMood.get().put(i + "", object.getInt("typeMood"));
                valueIdMood.get().put((i + 1) + "", object.getString("id"));
            }
            System.out.println(jsonArray.toString(2));
        } else {
            JSONObject jsonObject = new JSONObject(responseBody);
            System.out.println(jsonObject.toString(2));
        }
    }


    public static void writeToExcel() throws IOException {
        String pathExcel = "E:\\Testing\\Excel\\TypeMoodExcel.xlsx";
        FileInputStream fileStream = new FileInputStream(pathExcel);
        Workbook wb = new XSSFWorkbook(fileStream);
        Sheet sheet = wb.createSheet("Mood Data2");

        CellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.RED.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Row rowHeader = sheet.createRow(0);

        Cell cellID = rowHeader.createCell(0);
        cellID.setCellValue("ID Mood");
        cellID.setCellStyle(style);

        Cell cellValue = rowHeader.createCell(1);
        cellValue.setCellValue("VALUE");
        cellValue.setCellStyle(style);

        for (int i = 1; i < 19; i++) {
            Row rowValue = sheet.createRow(i);
            rowValue.createCell(0).setCellValue(valueIdMood.get().get("" + i));
            rowValue.createCell(1).setCellValue(classificationMood(valueMood.get().get("" + i)));
            System.out.println(valueIdMood.get().get("" + i));
        }

        FileOutputStream write = new FileOutputStream(pathExcel);
        wb.write(write);
        write.close();
        wb.close();
    }

    public static String classificationMood(int mood) {
        switch (mood) {
            case 1:
                return "Senang";
            case 2:
                return "Kecewa";
            default:
                return "Marah";
        }
    }

    public static void getTableValues() {
        // //table[contains(@class,"table-lg")]//tbody/tr[1]/td[4]
        List<WebElement> rowTable = driver.get().findElements(By.xpath("//table[contains(@class,\"table-lg\")]//tbody/tr"));
        String getMataUang = "";
        String getKursJual = "";
        String getKursBeli = "";
        for (int i = 1; i <= rowTable.size(); i++) {
            for (int j = 1; j <= 4; j++) {

                if (j == 1) {
                    getMataUang = driver.get().findElement(By.xpath("//table[contains(@class,\"table-lg\")]//tbody/tr[" + i + "]/td[" + j + "]")).getText();
                    mataUangWeb.get().put(i, getMataUang);
                } else if (j == 3) {
                    getKursJual = driver.get().findElement(By.xpath("//table[contains(@class,\"table-lg\")]//tbody/tr[" + i + "]/td[" + j + "]")).getText();
                    kursJualWeb.get().put(i, getKursJual);
                } else if (j == 4) {
                    getKursBeli = driver.get().findElement(By.xpath("//table[contains(@class,\"table-lg\")]//tbody/tr[" + i + "]/td[" + j + "]")).getText();
                    kursBeliWeb.get().put(i, getKursBeli);
                }
            }
        }
        System.out.println("mata uang: " + getMataUang + "| kur jual: " + getKursJual + " | kurs beli : " + getKursBeli);
    }

    public static void getAPIKurs() {
        Response response = given().baseUri("https://643eb0b4c72fda4a0bfe2c2c.mockapi.io").when().get("/kurstransaksi");

        String responseBody = response.asString();

        if (responseBody.trim().startsWith("[")) {
            JSONArray jsonArray = new JSONArray(responseBody);
            for (int i = 0; i < jsonArray.length(); i++) {
                JSONObject object = jsonArray.getJSONObject(i);
                mataUangAPI.get().put(i, object.getString("mata_uang"));
                kursJualAPI.get().put(i, object.getDouble("kurs_jual"));
                kursBeliAPI.get().put(i, object.getDouble("kurs_beli"));
            }
            System.out.println(jsonArray.toString(2));
        } else {
            JSONObject jsonObject = new JSONObject(responseBody);
            System.out.println(jsonObject.toString(2));
        }
    }

    public static void compareKurs() throws IOException {
        String pathExcel = "E:\\Testing\\Excel\\TypeMoodExcel.xlsx";
        FileInputStream fileStream = new FileInputStream(pathExcel);
        Workbook wb = new XSSFWorkbook(fileStream);
        Sheet sheet = wb.createSheet("Kurs");

        CellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.RED.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Row rowHeader = sheet.createRow(0);

        Cell cellHeaderMataUang = rowHeader.createCell(0);
        cellHeaderMataUang.setCellValue("MATA UANG");
        cellHeaderMataUang.setCellStyle(style);

        Cell cellHeaderKursBeliAPI = rowHeader.createCell(1);
        cellHeaderKursBeliAPI.setCellValue("KURS BELI API");
        cellHeaderKursBeliAPI.setCellStyle(style);

        Cell cellHeaderKursBeliWEB = rowHeader.createCell(2);
        cellHeaderKursBeliWEB.setCellValue("KURS BELI WEB");
        cellHeaderKursBeliWEB.setCellStyle(style);

        Cell cellHeaderKursJualAPI = rowHeader.createCell(3);
        cellHeaderKursJualAPI.setCellValue("KURS JUAL API");
        cellHeaderKursJualAPI.setCellStyle(style);

        Cell cellHeaderKursJualWEB = rowHeader.createCell(4);
        cellHeaderKursJualWEB.setCellValue("KURS JUAL WEB");
        cellHeaderKursJualWEB.setCellStyle(style);

        Cell cellHeaderHasilKursJual = rowHeader.createCell(5);
        cellHeaderHasilKursJual.setCellValue("HASIL KURS JUAL");
        cellHeaderHasilKursJual.setCellStyle(style);

        Cell cellHeaderHasilKursBeli = rowHeader.createCell(6);
        cellHeaderHasilKursBeli.setCellValue("HASIL KURS BELI");
        cellHeaderHasilKursBeli.setCellStyle(style);

        for (int i = 0; i < mataUangAPI.get().size(); i++) {
            for (int j = 1; j < mataUangWeb.get().size(); j++) {
                if (mataUangAPI.get().get(i).equals(mataUangWeb.get().get(j))) {
                    System.out.println(mataUangAPI.get().get(i) + mataUangWeb.get().get(j));
                    DecimalFormat df = (DecimalFormat) DecimalFormat.getInstance(Locale.GERMANY);

                    // Pastikan selalu 2 angka di belakang koma
                    df.applyPattern("#,###.00");
                    String kursBeliApi = df.format(kursBeliAPI.get().get(i));
                    String kursJualApi = df.format(kursJualAPI.get().get(i));
                    String hasilBeli = kursBeliApi.equals(kursBeliWeb.get().get(j))
                            ? "Accurate" : "Not Accurate";
                    String hasilJual = kursJualApi.equals(kursJualWeb.get().get(j))
                            ? "Accurate" : "Not Accurate";

                    Row rowValue = sheet.createRow(i+1);
                    rowValue.createCell(0).setCellValue(mataUangAPI.get().get(i));
                    rowValue.createCell(1).setCellValue(kursBeliApi);
                    rowValue.createCell(2).setCellValue(kursBeliWeb.get().get(j));
                    rowValue.createCell(3).setCellValue(kursJualApi);
                    rowValue.createCell(4).setCellValue(kursJualWeb.get().get(j));
                    rowValue.createCell(5).setCellValue(hasilJual);
                    rowValue.createCell(6).setCellValue(hasilBeli);
                }
            }
        }

        FileOutputStream write = new FileOutputStream(pathExcel);
        wb.write(write);
        write.close();
        wb.close();
    }

    public static void wait_for_second(Long duration) throws InterruptedException {
        Thread.sleep(duration);
    }

    public static void quit_browser() {
        driver.get().quit();
        driver.remove();
    }
}

