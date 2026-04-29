package org.example;

import com.sun.jna.platform.win32.User32;
import com.sun.jna.platform.win32.WinDef;
import io.restassured.http.ContentType;
import io.restassured.response.Response;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.Files;
import java.text.DecimalFormat;
import java.time.Duration;
import java.util.*;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

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

    public static int id;

    public static int amountFolder = 0;

    public static int index = 0;
    public static int size = 0;

    public static HashMap<String, Object> dataObject = new HashMap<>();
    public static HashMap<String, String> dataTemp = new HashMap<>();

    public static void main(String[] args) throws IOException, InterruptedException {

        System.setProperty("webdriver.chrome.driver", "C:\\Users\\Lenovo\\Downloads\\chromedriver-win64 (2)\\chromedriver-win64\\chromedriver.exe");
        driver.set(new ChromeDriver());
        driver.get().manage().window().maximize();
        driver.get().get("https://main.ocean.bca.co.id/visitor/product");
        wait_for_second(5000L);
        screenshot();
        scroll_to_element("//div[contains(text(),'Lihat Produk Lainnya')]");
        wait_for_second(2000L);
        clickElement("//div[contains(text(),'Lihat Produk Lainnya')]");
        wait_for_second(2000L);
        get_size_elements();
        for (int i = 1; i <= 2; i++) {
            String xpathCard = "(//div[contains(@id,'target')]//following-sibling::div[contains(@class,'w-full')]//div[contains(@class,'grid items-stretch')]//div[contains(@class,'rounded-xl')])" + "[" + i + "]";
//            TODO PAKAI ACUAN DARI XPATHCARD UNTUK GET XPATH TITLE AMBIL INDEXNYA DARI XPATH TITLE CTH (//div[contains(@id,'target')]//following-sibling::div[contains(@class,'w-full')]//div[contains(@class,'grid items-stretch')]//div[contains(@class,'rounded-xl')])"+"["+i+"]" //div[contains(@class,'text-base')];
//            TODO CTH = (//div[contains(@class,'rounded-xl')])[1]//div[contains(@class,'text-base')]
            String xpathTitle = "(//div[contains(@id,'target')]//following-sibling::div[contains(@class,'w-full')]//div[contains(@class,'grid items-stretch')]//div[contains(@class,'rounded-xl')]//div[contains(@class,'text-base')])[" + i + "]";

            scroll_to_element(xpathCard);
            wait_for_second(1000L);
            screenshot();
            String nameProduk = get_text(xpathTitle);
            String xpathBtn = "//div[contains(@id,'target')]//following-sibling::div[contains(@class,'w-full')]//div[contains(@class,'grid items-stretch')]//div[contains(@class,'rounded-xl')]//div[contains(text(),'" + nameProduk + "')]/following-sibling::a";

            String btnSelengkapnyaExist = check_exist_element(xpathBtn);
            String href = get_href(xpathBtn);
            String btnKembaliExist = "NOT EXIST";
            String nameProductDetails = "NOT EXIST";
            String status = "NOT FOUND";
            if (btnSelengkapnyaExist.equalsIgnoreCase("EXIST")) {
                clickElement(xpathBtn);
                status = check_status_browser();
                wait_for_second(4000L);
                switch_window("LAST");
                btnKembaliExist = check_exist_element("//p[normalize-space()='Kembali ke Produk']/parent::a");
                nameProductDetails = get_text("//a//following-sibling::h1");
                switch_window("FIRST");
            }

            dataTemp.put("NAMA PRODUK HALAMAN DEPAN", nameProduk);
            dataTemp.put("BUTTON SELENGKAPNYA (EXIST/NOT EXIST)", btnSelengkapnyaExist);
            dataTemp.put("URL", href);
            dataTemp.put("BUTTON KEMBALI (EXIST/NOT EXIST)", btnKembaliExist);
            dataTemp.put("NAMA PRODUK HALAMAN BELAKANG", nameProductDetails);
            dataTemp.put("STATUS", status);

            dataObject.put(String.valueOf(i), new HashMap<>(dataTemp));

//            System.out.println(i + " | " + nameProduk + " | " + btnSelengkapnyaExist + " | " + href);
        }

        System.out.println(dataObject);
        compare_data("NAMA PRODUK HALAMAN DEPAN,NAMA PRODUK HALAMAN BELAKANG");
        create_report();

//        TODO KETIKA ADA YANG TIDAK MEMILIKI BTN A (NOT EXIST MAKA AKAN MENGAMBIL DARI INDEX BERIKUTYA) XPATHNYA HARUS DIBENARKAN
//        ITEM PRODUK 1 TIDAK ADA BTN SELENGKAPNYA, TAPI KARENA MENGGUNAKAN INDEX MAKA AKAN MENGAMBIL DARI PRODUK 2 KARENA INDEX BTN SELENGKAPNYA ITU INDEX 1

    }


    //xpath
    // //div[contains(text(),'Lihat Produk Lainnya')] -> btn lihat produk lainnya
    // //div[contains(@id,'target')]//following-sibling::div[contains(@class,'w-full')]//div[contains(@class,'grid items-stretch')]//div[contains(@class,'rounded-xl')] -> card produk
    // //div[contains(@id,'target')]//following-sibling::div[contains(@class,'w-full')]//div[contains(@class,'grid items-stretch')]//div[contains(@class,'rounded-xl')]//div[contains(@class,'text-base')]
    // //div[contains(@id,'target')]//following-sibling::div[contains(@class,'w-full')]//div[contains(@class,'grid items-stretch')]//div[contains(@class,'rounded-xl')]//a -> btn href selengkapnya

    //TODO SWITCH WINDOW
    public static void switch_window(String index) {
        Set<String> handles = driver.get().getWindowHandles();

        // convert ke array/list biar bisa pakai index
        String[] tabs = handles.toArray(new String[0]);

        if (index.equalsIgnoreCase("FIRST")) {
            driver.get().close();
        }

        // pindah ke tab kedua
        driver.get().switchTo().window(tabs[index.equalsIgnoreCase("FIRST") ? 0 : 1]);


    }

    //TODO CHECK WINDOW (BLANK/404)
    public static String check_status_browser() {
        String title = driver.get().getTitle();
        String status = "";
        if (title.toLowerCase().contains("404") || title.toLowerCase().contains("not found")) {
            status = "FAILED";
        } else {
            status = "PASSED";
        }
        return status;
    }

    //TODO GET SIZE ELEMENT
    public static void get_size_elements() {
        size = driver.get().findElements(By.xpath("//div[contains(@id,'target')]//following-sibling::div[contains(@class,'w-full')]//div[contains(@class,'grid items-stretch')]//div[contains(@class,'rounded-xl')]")).size();
    }

    //TODO SCROLL TO ELEMENT BY INDEX
    public static void scroll_to_element(String xpath) throws InterruptedException {
        WebElement element = driver.get().findElement(By.xpath(xpath));

        JavascriptExecutor js = (JavascriptExecutor) driver.get();
        js.executeScript("arguments[0].scrollIntoView(true);", element);

        Thread.sleep(1000);
    }

    //TODO GET TEXT ELEMENT BY INDEX
    public static String get_text(String xpath) {
        if (driver.get().findElements(By.xpath(xpath)).size() > 0) {
            return driver.get().findElement(By.xpath(xpath)).getText();
        } else {
            return "NOT EXIST";
        }
    }

    //TODO CHECK ELEMENT EXIST BY INDEX
    public static String check_exist_element(String xpath) {
        if (driver.get().findElements(By.xpath(xpath)).size() > 0) {
            return "EXIST";
        } else {
            return "NOT EXIST";
        }
    }


    //TODO GET HREF ELEMENT
    public static String get_href(String xpath) {
        try {
            WebDriverWait wait = new WebDriverWait(driver.get(), Duration.ofSeconds(5));

            WebElement link = wait.until(
                    ExpectedConditions.presenceOfElementLocated(By.xpath(xpath))
            );

            String href = link.getAttribute("href");
            return (href != null && !href.isEmpty()) ? href : "NOT FOUND";

        } catch (TimeoutException e) {
            return "NOT FOUND";
        }
    }

    //TODO HEADER LIST
    public static ArrayList<String> arrHeader = new ArrayList<String>(
            Arrays.asList("NO", "NAMA PRODUK HALAMAN DEPAN", "NAMA PRODUK HALAMAN BELAKANG", "BUTTON SELENGKAPNYA (EXIST/NOT EXIST)", "BUTTON KEMBALI (EXIST/NOT EXIST)", "URL", "STATUS")
    );

    //TODO METHOD COMPARE (FUNGSINYA UNTUK PERINTAH FIELD MANA AJA YANG INGIN DICOMPARE, BASED ON HEADER)
    //TODO FORMATNYA NAMA1,NAMA2;DLL
    public static ArrayList<String> dataCompare = new ArrayList<>();

    public static void compare_data(String str) {
        String[] splitData = str.split(";");
        dataCompare.addAll(Arrays.asList(splitData));
    }

    //TODO CREATE REPORT
    public static void create_report() throws IOException {
        String pathExcel = "E:\\TEST\\Report.xlsx";

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("REPORT VERIFY");

        CellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.RED.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Row rowHeader = sheet.createRow(0);

        for (int i = 0; i < arrHeader.size(); i++) {
            Cell cellID = rowHeader.createCell(i);
            cellID.setCellValue(arrHeader.get(i));
            cellID.setCellStyle(style);

        }

        for (int j = 1; j <= dataObject.size(); j++) {
            HashMap<String, String> data = (HashMap<String, String>) dataObject.get(String.valueOf(j));

            HashMap<String, String> resultCompare = new HashMap<>();
            if (dataCompare.size() > 0) {
                for (String val : dataCompare) {
                    String[] split = val.split(",");
                    String dataA = split[0];
                    String dataB = split[1];
                    if (data.get(dataA).equalsIgnoreCase(data.get(dataB))) {
                        resultCompare.put(dataA, "MATCH");
                        resultCompare.put(dataB, "MATCH");
                    } else {
                        resultCompare.put(dataA, "NOT MATCH");
                        resultCompare.put(dataB, "NOT MATCH");
                    }
                }
            }

            //write number
            Row rowValue = sheet.createRow(j);
            rowValue.createCell(0).setCellValue(j);

            CellStyle styleMatch = wb.createCellStyle();
            styleMatch.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
            styleMatch.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            CellStyle styleNotMatch = wb.createCellStyle();
            styleNotMatch.setFillForegroundColor(IndexedColors.RED.getIndex());
            styleNotMatch.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            for (int k = 0; k < data.keySet().size(); k++) {
                Cell cell = rowValue.createCell(k + 1);
                cell.setCellValue(data.get(arrHeader.get(k + 1)));
                if (resultCompare.containsKey(arrHeader.get(k+1))) {

                    String status = resultCompare.get(arrHeader.get(k+1));

                    if ("MATCH".equalsIgnoreCase(status)) {
                        cell.setCellStyle(styleMatch);
                    } else {
                        cell.setCellStyle(styleNotMatch);
                    }
                }
            }

            resultCompare.clear();
        }

        FileOutputStream write = new FileOutputStream(pathExcel);
        wb.write(write);
        write.close();
        wb.close();
    }

    public static void taruhFile(String path, String name) throws IOException {

        File parent = new File("E:\\ANIM\\F1");

        File[] childs = parent.listFiles(File::isDirectory);

        if (childs == null || childs.length == 0) {
            System.out.println("Child folder tidak ditemukan");
            //ketika childnya cuman satu maka pakai ini saja
        } else if (childs.length > 1) {
            // ambil child terakhir
            File latestFolder = Arrays.stream(childs)
                    .min(Comparator.comparingLong(File::lastModified))
                    .orElse(null);

            File source = new File(path);

            File target = new File(latestFolder, name);

            FileUtils.copyFile(source, target);

        }


    }

    private static void clickElement(String xpath) throws InterruptedException {
        WebElement element = driver.get().findElement(By.xpath(xpath));
        element.click();
        wait_for_second(2000L);
    }


    private static void postTestCase() {
        Map<String, Object> testCasePayload = new HashMap<>();
        testCasePayload.put("name", "Login Test Scenario");

        Response response = given()
                .contentType(ContentType.JSON)
                .body(testCasePayload)
                .when()
                .post("http://127.0.0.1:5000/api/tests")
                .then()
                .statusCode(201)
                .extract().response();

// Ambil ID dari response untuk digunakan saat upload screenshot
        int testId = response.path("id");
        id = testId;
        System.out.println("Test Case Created with ID: " + testId);
    }

    public static void screenshot() throws IOException {
        // Ambil screenshot
        File src = ((TakesScreenshot) driver.get()).getScreenshotAs(OutputType.FILE);

        String currentTime = String.valueOf(System.currentTimeMillis());
        String fileName = "E:/PROJECTS/screenshots/screenshot_" + currentTime + ".png";

        // Simpan ke local
//        File dest = new File(fileName);
//        FileUtils.copyFile(src, dest);

//        System.out.println(convertFileToBase64(dest));
        long startTime = System.currentTimeMillis();
//        Map<String, Object> body = new HashMap<>();
//        body.put("step_name", "screenshot_" + currentTime + ".png");
//        body.put("image_base64", getResizedBase64(src, 1400)); // String panjang hasil encode
//
//        given()
//                .contentType(ContentType.JSON)
//                .body(body)
//                .when()
//                .post("http://127.0.0.1:5000/api/test/" + id + "/screenshots")
//                .then();
    }

    public static String getResizedBase64(File srcFile, int targetWidth) throws IOException {
        // 1. Baca file gambar mentah
        BufferedImage originalImage = ImageIO.read(srcFile);

        // 2. Hitung tinggi proporsional agar gambar tidak gepeng
        double aspectRatio = (double) originalImage.getHeight() / originalImage.getWidth();
        int targetHeight = (int) (targetWidth * aspectRatio);

        // 3. Buat kanvas baru dengan ukuran yang dikecilkan
        BufferedImage resizedImage = new BufferedImage(targetWidth, targetHeight, BufferedImage.TYPE_INT_RGB);

        // 4. Gambar ulang gambar asli ke kanvas baru (proses resize)
        Graphics2D g2d = resizedImage.createGraphics();
        // Menggunakan rendering hints untuk kualitas resize yang baik
        g2d.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
        g2d.drawImage(originalImage, 0, 0, targetWidth, targetHeight, null);
        g2d.dispose();

        // 5. Konversi langsung dari memori ke Base64 (tanpa simpan file)
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        // Gunakan "jpg" untuk ukuran file yang jauh lebih kecil dibanding "png"
        ImageIO.write(resizedImage, "jpg", baos);

        return Base64.getEncoder().encodeToString(baos.toByteArray());
    }

    public static String convertFileToBase64(File imageFile) throws IOException {
        // Baca file menjadi byte array
        byte[] fileContent = Files.readAllBytes(imageFile.toPath());

        // Encode byte array menjadi String Base64
        return Base64.getEncoder().encodeToString(fileContent);
    }


    public static void open_multiple_file(File file) throws IOException, InterruptedException {
        if (file.getName().contains(".txt")) {
            Runtime.getRuntime().exec("notepad " + file.getAbsolutePath());
        } else if (file.getName().toLowerCase().endsWith(".png") || file.getName().toLowerCase().endsWith(".jpg")) {
            Runtime.getRuntime().exec("cmd /c start \"\" \"" + file.getAbsoluteFile() + "\"");
            Thread.sleep(3000);
        } else {
            Runtime.getRuntime().exec("cmd /c start \"\" \"" + file.getAbsoluteFile() + "\"");
            Thread.sleep(3000);
        }

        Thread.sleep(1200); // tunggu notepad terbuka

        maximaze_file();

        close_file();
    }

    public static void maximaze_file() throws InterruptedException {
        WinDef.HWND hwnd = User32.INSTANCE.GetForegroundWindow();

        if (hwnd != null) {

            User32.INSTANCE.ShowWindow(hwnd, 3); // maximize

            System.out.println("Notepad maximize");
        }
    }

    public static void close_file() throws InterruptedException {
//        WinDef.HWND hwnd = User32.INSTANCE.FindWindow("Notepad", null);

        WinDef.HWND hwnd = User32.INSTANCE.GetForegroundWindow();
        Thread.sleep(1000);

        if (hwnd != null) {

            // close
            User32.INSTANCE.PostMessage(hwnd, 0x0010, null, null);

        } else {
            System.out.println("Notepad tidak ditemukan");
        }
    }

    public static void openFolders(File directory) throws IOException, InterruptedException {

        if (!directory.exists()) return;

//        new ProcessBuilder("explorer", directory.getAbsolutePath()).start();
//        Runtime.getRuntime().exec(
//                "cmd /c start \"\" \"" + directory.getAbsolutePath() + "\""
//        );

        Desktop.getDesktop().open(directory);
        Thread.sleep(1000);
        maximaze_file();
        Thread.sleep(1000);
        close_file();

        File[] files = directory.listFiles();

        if (files != null) {
            for (File f : files) {
                if (f.isDirectory()) {
                    openFolders(f);
                }
            }
        }
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

                    Row rowValue = sheet.createRow(i + 1);
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

    public static File get_path_sub_folder(File dir) {
        File[] files = dir.listFiles();
        if (files == null) return null;

        for (File file : files) {
            if (file.isDirectory()) {
                File subResult = get_path_sub_folder(file);

                // Jika di dalamnya masih ada folder, kembalikan hasil terdalam
                if (subResult != null) {
                    return subResult;
                }

                // Jika tidak ada subfolder lagi, kembalikan folder ini
                return file;
            }
        }
        return null;
    }

    public static void copy_file(File dir) throws IOException {
        FileUtils.copyFile(new File("E:\\KKN\\VIDEO\\ACTIVITY 1\\ASSET\\IMG_20231130_091408.jpg"), new File("E:\\ANIM\\F1\\F2\\F3\\hilmi.jpg"));
    }

    public static void copyToAllFolders(File rootDir) throws InterruptedException {

        File sourceFile = new File("E:\\KKN\\VIDEO\\ACTIVITY 1\\ASSET\\IMG_20231130_091408.jpg");

        File[] folders = rootDir.listFiles(File::isDirectory);

        ExecutorService executor = Executors.newFixedThreadPool(5);

        for (File folder : folders) {

            executor.submit(() -> {
                try {

                    File targetFolder = get_path_sub_folder(folder);

                    if (targetFolder != null) {
                        File targetFile = new File(targetFolder, "hilmi.jpg");

                        FileUtils.copyFile(sourceFile, targetFile);

                        System.out.println("Copied to: " + targetFile.getAbsolutePath());
                    }

                } catch (Exception e) {
                    e.printStackTrace();
                }
            });
        }

        executor.shutdown();
    }
}

