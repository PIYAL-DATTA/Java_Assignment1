/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package javaassignmnet1;

/**
 *
 * @author 19020
 */
import java.time.DayOfWeek;
import java.time.LocalDate; // import the LocalDate class
import java.io.FileInputStream; // Apache POI
import java.io.FileNotFoundException;  // Import this class to handle errors
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class JavaAssignmnet1 {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws FileNotFoundException, EncryptedDocumentException, IOException, InvalidFormatException, InterruptedException {

        // TODO code application logic here
        LocalDate myObj = LocalDate.now(); // Create a date object
        System.out.println(myObj); // Display the current date
        DayOfWeek day = myObj.getDayOfWeek();
        System.out.println(day); // Display the current day of the week
        String sheetname;
        if ("SATURDAY".equals(day.toString())) {
            sheetname = "Saturday";
        } else if ("SUNDAY".equals(day.toString())) {
            sheetname = "Sunday";
        } else if ("MONDAY".equals(day.toString())) {
            sheetname = "Monday";
        } else if ("TUESDAY".equals(day.toString())) {
            sheetname = "Tuesday";
        } else if ("WEDNESDAY".equals(day.toString())) {
            sheetname = "Wednesday";
        } else if ("THURSDAY".equals(day.toString())) {
            sheetname = "Thursday";
        } else {
            sheetname = "Friday";
        }
        // EXCEL file read
        FileInputStream loc = new FileInputStream("G:\\java\\JavaAssignmnet1\\4BeatsQ1.xlsx");
        Workbook wb = WorkbookFactory.create(loc);
        Sheet sheet = wb.getSheet(sheetname); // Sheet based on days
        
        // Search value read
        for (int i = 2; i < 12; i++) {
            Row row = sheet.getRow(i);
            Cell cell = row.getCell(2);
            String data = cell.getStringCellValue();
            
            // Checking Empty DATA
            if(data.isEmpty()){
                System.out.println("Empty value.");
                continue;
            }
            
            // Chromedriver setup
            System.setProperty("webdriver.chrome.driver", "C:\\chromedriver-win64\\chromedriver.exe");
            WebDriver driver = new ChromeDriver();
            driver.get("https://www.google.com/");
            driver.manage().window().maximize();
            
            // Wait until the search box is present
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
            WebElement searchBox = wait.until(ExpectedConditions.presenceOfElementLocated(By.name("q")));
            searchBox.sendKeys(data); // Sending search value
            Thread.sleep(1000);
            
            // For search result list
            List<WebElement> list = driver.findElements(By.xpath("//ul[@role='listbox']/li"));
            System.out.println("Searched Elements: " + data);
            int maxlength = 0, minlength = Integer.MAX_VALUE;
            String max = "null", min = "null";
            for (WebElement element : list) {
                if(element.getText().length() > maxlength){
                    maxlength = element.getText().length();
                    max = element.getText();
                }
                if(element.getText().length() < minlength){
                    minlength = element.getText().length();
                    min = element.getText();
                }
            }
            System.out.println(max + " length: " + maxlength);
            System.out.println(min + " length: " + minlength);
            
            // For inserting values in Excel
            Cell maxcell = row.createCell(3); // Max Value
            maxcell.setCellValue(max);
            FileOutputStream outputmax = new FileOutputStream("G:\\java\\JavaAssignmnet1\\4BeatsQ1.xlsx");
            wb.write(outputmax);
            Cell mincell = row.createCell(4); // Min Value
            mincell.setCellValue(min);
            FileOutputStream outputmin = new FileOutputStream("G:\\java\\JavaAssignmnet1\\4BeatsQ1.xlsx");
            wb.write(outputmin);
            driver.quit();
        }
    }
}