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
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static io.restassured.RestAssured.given;

public class Main {

    public static ThreadLocal<WebDriver> driver = new ThreadLocal<>();
    public static ThreadLocal<Map<String, String>> valueIdMood = ThreadLocal.withInitial(() -> new HashMap<>());
    public static ThreadLocal<Map<String, Integer>> valueMood = ThreadLocal.withInitial(() -> new HashMap<>());

    public static ThreadLocal<Map<String, String>> mataUangWeb = ThreadLocal.withInitial(() -> new HashMap<>());

    public static ThreadLocal<Map<String, Double>> kursJualWeb = ThreadLocal.withInitial(() -> new HashMap<>());

    public static ThreadLocal<Map<String, Double>> kursBeliWeb = ThreadLocal.withInitial(() -> new HashMap<>());

    public static void main(String[] args) throws IOException, InterruptedException {
        System.setProperty("webdriver.chrome.driver", "C:\\Users\\Lenovo\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");
        driver.set(new ChromeDriver());
        driver.get().manage().window().maximize();
        driver.get().get("https://www.bi.go.id/id/statistik/informasi-kurs/transaksi-bi/default.aspx");
        wait_for_second(2000L);
        getTableValues();
        quit_browser();
//        getAPI();
//        writeToExcel();

    }

    public static void getAPI() {
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
                } else if (j == 3) {
                    getKursJual = driver.get().findElement(By.xpath("//table[contains(@class,\"table-lg\")]//tbody/tr[" + i + "]/td[" + j + "]")).getText();
                } else if (j == 4) {
                    getKursBeli = driver.get().findElement(By.xpath("//table[contains(@class,\"table-lg\")]//tbody/tr[" + i + "]/td[" + j + "]")).getText();
                }
            }
        }
        System.out.println("mata uang: " + getMataUang + "| kur jual: " + getKursJual + " | kurs beli : " + getKursBeli);
    }

    public static void wait_for_second(Long duration) throws InterruptedException {
        Thread.sleep(duration);
    }

    public static void quit_browser() {
        driver.get().quit();
        driver.remove();
    }
}

