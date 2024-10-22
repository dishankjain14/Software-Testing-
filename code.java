package project ;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;
import java.time.Duration;
import org.apache.poi.ss.usermodel.DataFormatter;

public class code {
    WebDriver driver;
    Workbook workbook;
    Sheet sheet;

    @BeforeClass
    public void setup() throws IOException {
        // Set up Chrome WebDriver using WebDriverManager
        io.github.bonigarcia.wdm.WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();

        // Maximize the window to full-screen
        driver.manage().window().maximize();

        // Load the Excel file
        FileInputStream fis = new FileInputStream("C:\\Users\\disha\\eclipse-workspace\\softwaretesting\\target\\inputs.xlsx");
        workbook = new XSSFWorkbook(fis);
        sheet = workbook.getSheetAt(0); // Assuming your data is in the first sheet
    }

    @Test
    public void testFormFilling() throws IOException {
        DataFormatter dataFormatter = new DataFormatter(); // Create a DataFormatter

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);

            // Skip row based on "Control" column value
            Cell controlCell = row.getCell(0);
            String control = controlCell.getStringCellValue();
            if (control.equalsIgnoreCase("No")) {
                continue;
            }

            // Open the test webpage
            driver.get("https://www.tutorialspoint.com/selenium/practice/selenium_automation_practice.php");

            // Fill out the form fields in the specified order
            driver.findElement(By.name("name")).sendKeys(dataFormatter.formatCellValue(row.getCell(1))); // First Name
            driver.findElement(By.name("email")).sendKeys(dataFormatter.formatCellValue(row.getCell(2))); // Email

            // Gender
            Cell genderCell = row.getCell(3); // Gender is at index 3
            String gender = genderCell != null ? dataFormatter.formatCellValue(genderCell) : null;

            // Handle Gender selection
            if (gender != null) {
                if (gender.equalsIgnoreCase("Male")) {
                    driver.findElement(By.xpath("//input[@name='gender' and @type='radio'][1]")).click(); // Male
                } else if (gender.equalsIgnoreCase("Female")) {
                    driver.findElement(By.xpath("//label[text()='Female']/preceding-sibling::input")).click(); // Female
                } else if (gender.equalsIgnoreCase("Other")) {
                    driver.findElement(By.xpath("//label[text()='Other']/preceding-sibling::input")).click(); // Other
                } else {
                    System.out.println("Gender value not recognized: " + gender);
                }
            } else {
                System.out.println("Gender value is null, cannot select gender.");
            }


            // Mobile
            Cell mobileCell = row.getCell(4); // Mobile is at index 4
            String mobile = mobileCell != null ? dataFormatter.formatCellValue(mobileCell) : null;

            // Fill in the mobile number in the form
            if (mobile != null) {
                driver.findElement(By.name("mobile")).sendKeys(mobile); // Use the formatted string value
            } else {
                System.out.println("Mobile number is null, cannot fill the form.");
            }
            // Date of Birth
            String dob = dataFormatter.formatCellValue(row.getCell(5)); // Get the DOB string (e.g., "08-09-2003")

            // Convert to the format YYYY-MM-DD
            String[] dobParts = dob.split("-");
            if (dobParts.length == 3) {
                String day = dobParts[0];
                String month = dobParts[1];
                String year = dobParts[2];

                // Format the date as YYYY-MM-DD
                String formattedDOB = year + "-" + month + "-" + day;

                // Fill in the date of birth in the form
                driver.findElement(By.name("dob")).sendKeys(formattedDOB); // Enter the formatted DOB
            } else {
                System.out.println("DOB format is incorrect: " + dob);
            }
            // Subjects
            driver.findElement(By.name("subjects")).sendKeys(dataFormatter.formatCellValue(row.getCell(6))); // Subjects

            // Handle Hobbies
            String hobbies = dataFormatter.formatCellValue(row.getCell(7)); // Assuming hobbies are in column 7
            if (hobbies.contains("Sports")) {
                driver.findElement(By.xpath("//label[text()='Sports']/preceding-sibling::input")).click(); // Sports
            }
            if (hobbies.contains("Reading")) {
                driver.findElement(By.xpath("//label[text()='Reading']/preceding-sibling::input")).click(); // Reading
            }
            if (hobbies.contains("Music")) {
                driver.findElement(By.xpath("//label[text()='Music']/preceding-sibling::input")).click(); // Music
            }

            // Profile Photo upload
            WebElement photoUpload = driver.findElement(By.name("picture")); // Updated name to "file-553"
            String photoPath = dataFormatter.formatCellValue(row.getCell(8)); // Assuming the photo file path is in column 8
            photoUpload.sendKeys(photoPath); // Sends the file path to the input field for upload

            // Fill the current address
            driver.findElement(By.xpath("/html/body/main/div/div/div[2]/form/div[9]/div/textarea")).sendKeys(dataFormatter.formatCellValue(row.getCell(9))); // Current Address

            // Select State
            String state = row.getCell(10).getStringCellValue(); // Assuming state is in column 10
            WebElement stateDropdown = driver.findElement(By.id("state")); // Change from name to id
            Select selectState = new Select(stateDropdown);
            selectState.selectByVisibleText(state);

            // Select City
            String city = row.getCell(11).getStringCellValue(); // Assuming city is in column 11
            WebElement cityDropdown = driver.findElement(By.id("city")); // Change from name to id
            Select selectCity = new Select(cityDropdown);
            selectCity.selectByVisibleText(city);

            // Optionally submit the form
            // driver.findElement(By.id("submit-button-id")).click();

            // Add test result to Excel
            Cell resultCell = row.createCell(12);
            resultCell.setCellValue("Form submitted successfully!");
            
            
            // Save the updated Excel file
            FileOutputStream fos = new FileOutputStream("C:\\Users\\disha\\eclipse-workspace\\softwaretesting\\target\\outputs.xlsx");
            workbook.write(fos);
            fos.close();
        }
    }

    @AfterClass
    public void tearDown() throws IOException {
        // Close the WebDriver and Excel workbook
        driver.quit();
        workbook.close();
    }
}