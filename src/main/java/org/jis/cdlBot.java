package org.jis;

import org.apache.commons.exec.ExecuteWatchdog;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import javax.swing.*;
import java.io.*;
import java.time.Duration;
import java.time.LocalDate;

public class cdlBot {
    public static void main(String[] args) throws IOException, InterruptedException {
        //instantiate GUI
        cdlForm form = new cdlForm();
        //instantiate formatter to format name cell
        DataFormatter formatter = new DataFormatter();
        //instantiate current date
        LocalDate date = LocalDate.now();

        //create variables
        FileInputStream inputStream = null;
        FileInputStream inputStream2 = null;
        XSSFWorkbook wb = null;
        XSSFWorkbook wb2 = null;
        XSSFSheet sheet = null;
        XSSFSheet sheet2 = null;
        int rowCount = 0;
        int lastRow = 0;

        //start GUI
        form.run();
        form.setTxt1("Checking Files...");

        //set filepaths
        File file = new File("S:\\IT\\Projects\\CDL\\ReviewOrders.xlsx");
        File file2 = new File("S:\\IT\\Projects\\CDL\\CDLIssueReferenceNumbers.xlsx");

        //open files and get sheet count, then close
        FileInputStream inputStream0 = new FileInputStream(file);
        XSSFWorkbook wb0 = new XSSFWorkbook(inputStream0);
        int sheetNum = wb0.getNumberOfSheets();
        wb0.close();
        inputStream0.close();

        //if only 1 sheet run powershell script to format daily report, if more than 1 sheet than script has already been run
        try {
            if (sheetNum == 1) {
                form.setTxt1("Starting Powershell");
                String command = "powershell.exe S:\\IT\\Scripts\\CDL_Daily.ps1";
                Process powerShellProcess = Runtime.getRuntime().exec(command);
                BufferedReader stdInput1 = new BufferedReader(new InputStreamReader(powerShellProcess.getErrorStream()));
                String line1 = stdInput1.readLine();
                if(line1 != null){
                    System.out.println(line1);
                    JOptionPane.showMessageDialog(null,line1 + System.lineSeparator());
                }
                stdInput1.close();
                powerShellProcess.getOutputStream().close();
                powerShellProcess.waitFor();
                form.setTxt1("Finishing Powershell");
            }
        }catch(Exception e){
            JOptionPane.showMessageDialog(null, e.toString() );
        }

        form.setTxt1("Opening Files...");
        //open files and set needed variables
        try {
            inputStream = new FileInputStream(file);
            inputStream2 = new FileInputStream(file2);
            wb = new XSSFWorkbook(inputStream);
            wb2 = new XSSFWorkbook(inputStream2);
            sheet = wb.getSheet("Sixth");
            sheet2 = wb2.getSheet("Sheet2");
            rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
            lastRow = sheet2.getLastRowNum() + 1;
        }catch(Exception e) {
            JOptionPane.showMessageDialog(null, e.toString());
        }


        form.setTxt1("Starting WebDriver");
        form.setTxt2("Total Orders: " + rowCount);

        //start headless driver for chrome
        System.setProperty("webdriver.chrome.driver", "S:\\IT\\Projects\\CDL\\chromedriver.exe");
        ChromeOptions headless = new ChromeOptions();
        headless.addArguments("headless");
        WebDriver driver = new ChromeDriver(headless);
        //WebDriver driver = new ChromeDriver();
        //driver.manage().window().maximize();
        //if fails to find order id it will take a screenshot, this ensures whole page is captured
        driver.manage().window().setSize(new Dimension(1920,1080));
        driver.get("http://ship.cdldelivers.com/xcelerator/xagent/Index.aspx");

        form.setTxt1("Signing In");

        //fill elements with login info
        WebElement username = driver.findElement(By.name("AgentNo"));
        username.sendKeys("1550");
        WebElement password = driver.findElement(By.name("Password"));
        password.sendKeys("Cxt123!");
        WebElement login = driver.findElement(By.name("Submit"));
        login.click();

        form.setTxt1("Opening Orders");

        //wait until expected elements is present to proceed
        WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(60));
        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[text()[contains(.,'Open Orders (')]]")));

        //selects the open order option
        WebElement openOrder = driver.findElement(By.xpath("//*[text()[contains(.,'Open Orders (')]]"));
        openOrder.click();

        //run loop until i = rowcount
        for (int i = 1; i <= rowCount; i++) {
            cdlForm.setTxt1("Finding Order");
            cdlForm.setTxt2(i + "/" + rowCount);
            System.out.println("Row Number: " + i);
            //set path for screenshot if entry fails
            File failImage = new File("S:\\IT\\Projects\\CDL\\FailImage\\failRow" + i + "_" + date + ".png");

            //get values stored in specified cells
            double orderID = sheet.getRow(i).getCell(0).getNumericCellValue();
            String podDate = sheet.getRow(i).getCell(1).getStringCellValue();
            String podNameValue = formatter.formatCellValue(sheet.getRow(i).getCell(2));
            String pArrivalValue =  sheet.getRow(i).getCell(3).getStringCellValue() + " 04:00";
            String pDepartureValue =  sheet.getRow(i).getCell(3).getStringCellValue() + " 04:05";
            String dArrivalValue =  sheet.getRow(i).getCell(3).getStringCellValue() + " 04:10";
            String dDepartureValue =  sheet.getRow(i).getCell(3).getStringCellValue() + " 04:15";
            XSSFCell passFailCell = sheet.getRow(i).getCell(3);
            XSSFCell podNameCell = sheet.getRow(i).getCell(6);

            try {
                Thread.sleep(8000);
            } catch (InterruptedException ie) {
            }

            //find location of element using orderid
            boolean isPresent = driver.findElements(By.xpath("((//tr[.//b[text() = '" + orderID + "']]) [4])//input[@name='PickupArrival']")).size() > 0;

            //if orderid location is found
            if (isPresent) {
                //create cell's to store entered data for confirmation
                XSSFCell orderIDCell = sheet.getRow(i).createCell(4);
                XSSFCell dateCell = sheet.getRow(i).createCell(5);
                XSSFCell nameCell = sheet.getRow(i).createCell(6);
                XSSFCell paCell = sheet.getRow(i).createCell(7);
                XSSFCell pdCell = sheet.getRow(i).createCell(8);
                XSSFCell daCell = sheet.getRow(i).createCell(9);
                XSSFCell ddCell = sheet.getRow(i).createCell(10);

                form.setTxt1("Entering Data");
                form.setTxt3("Order ID Found");

                //find various elements and enter related data if empty
                WebElement pArrival = driver.findElement(By.xpath("((//tr[.//b[text() = '" + orderID + "']]) [4])//input[@name='PickupArrival']"));
                if (pArrival.getAttribute("value").isEmpty()) {
                    pArrival.sendKeys(pArrivalValue);
                }

                WebElement pDeparture = driver.findElement(By.xpath("((//tr[.//b[text() = '" + orderID + "']]) [4])//input[@name='PickupDeparture']"));
                if (pDeparture.getAttribute("value").isEmpty()) {
                    pDeparture.sendKeys(pDepartureValue);
                }

                WebElement dArrival = driver.findElement(By.xpath("((//tr[.//b[text() = '" + orderID + "']]) [4])//input[@name='DeliveryArrival']"));
                if (dArrival.getAttribute("value").isEmpty()) {
                    dArrival.sendKeys(dArrivalValue);
                }

                WebElement dDeparture = driver.findElement(By.xpath("((//tr[.//b[text() = '" + orderID + "']]) [4])//input[@name='DeliveryDeparture']"));
                if (dDeparture.getAttribute("value").isEmpty()) {
                    dDeparture.sendKeys(dDepartureValue);
                }

                WebElement podName = driver.findElement(By.xpath("((//tr[.//b[text() = '" + orderID + "']]) [4])//input[@name='PODname']"));
                podName.clear();
                podName.sendKeys(podNameValue);

                WebElement podCompletion = driver.findElement(By.xpath("((//tr[.//b[text() = '" + orderID + "']]) [4])//input[@name='PODcompletion']"));
                podCompletion.clear();
                podCompletion.sendKeys(podDate);

                //find and click save element
                WebElement save = driver.findElement(By.xpath("((//tr[.//b[text() = '" + orderID + "']]) [4])//input[@name='Save']"));
                save.click();

                //set cells with value entered into specified element
                passFailCell.setCellValue("Pass");
                orderIDCell.setCellValue(orderID);
                dateCell.setCellValue(podDate);
                nameCell.setCellValue(podNameValue);
                paCell.setCellValue(pArrivalValue);
                pdCell.setCellValue(pDepartureValue);
                daCell.setCellValue(dArrivalValue);
                ddCell.setCellValue(dDepartureValue);

                //open outputstream and write to file, then close stream
                FileOutputStream outputStream = new FileOutputStream(file);
                wb.write(outputStream);
                outputStream.close();

                //refresh webpage every 20 rows
                System.out.println("Order ID Found");
                if(i  % 20 == 0) {
                    driver.navigate().refresh();
                }
            }
            //if orderid not found in table
            else {
                //if podnamecell has a value tell user it is a duplicate entry
                if(podNameCell != null) {
                    form.setTxt3("Duplicate Order ID");
                }
                //otherwise failed to find orderid
                else {
                    form.setTxt3("Order ID Not Found");
                    //write to file that it failed
                    passFailCell.setCellValue("Fail");
                    FileOutputStream outputStream1 = new FileOutputStream(file);
                    wb.write(outputStream1);
                    outputStream1.close();

                    //set variables to hold the relevant info of the failed row, and format row
                    CellStyle cs = wb2.createCellStyle();
                    CellStyle cs2 = wb2.createCellStyle();
                    CreationHelper createHelper = wb2.getCreationHelper();

                    cs.setAlignment(HorizontalAlignment.LEFT);
                    cs2.setAlignment(HorizontalAlignment.LEFT);
                    cs2.setDataFormat(createHelper.createDataFormat().getFormat("mm/dd/yyyy"));

                    XSSFRow  newRow      = sheet2.createRow(lastRow);
                    XSSFCell failOrderID = newRow.createCell(0);
                    XSSFCell failPODDate = newRow.createCell(1);
                    XSSFCell failPODName = newRow.createCell(2);
                    XSSFCell failDate =    newRow.createCell(3);

                    failOrderID.setCellValue(orderID);
                    failOrderID.setCellStyle(cs);

                    failPODDate.setCellValue(podDate);
                    failPODDate.setCellStyle(cs);

                    failPODName.setCellValue(podNameValue);
                    failPODName.setCellStyle(cs);

                    failDate.setCellValue(date);
                    failDate.setCellStyle(cs2);

                    //open outputstream and write info to file 2
                    FileOutputStream outputStreamF = new FileOutputStream(file2);
                    wb2.write(outputStreamF);
                    outputStreamF.close();
                    lastRow++;
                    if(i  % 20 == 0) {
                        driver.navigate().refresh();
                    }

                    //take a screenshot of the website when orderid is unable to be found
                    File screenshot = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
                    try{
                        FileUtils.copyFile(screenshot, failImage);
                    }catch (IOException e){
                        JOptionPane.showMessageDialog(null, e.toString());
                    }
                }
            }
        }
        //open outputstream for file 1 to formate sheet after all data has been entered
        FileOutputStream outputStream2 = new FileOutputStream(file);
        //autosize columns
        for(int i = 4;i <= 10;i++) {
            sheet.autoSizeColumn(i);
        }

        wb.write(outputStream2);

        //close workbooks, outputstreams and driver
        wb.close();
        wb2.close();
        outputStream2.close();
        inputStream.close();
        inputStream2.close();
        driver.quit();

        form.setTxt1("Emailing File.");

        //start script to send email after all data has been entered
        System.out.println("ps start");
        String command2 = "powershell.exe S:\\IT\\Scripts\\CDL_Email.ps1";
        ExecuteWatchdog watchdog = new ExecuteWatchdog(20000);
        Process powerShellProcess2 = Runtime.getRuntime().exec(command2);
        if(watchdog != null){
            watchdog.start(powerShellProcess2);
        }
        BufferedReader stdInput = new BufferedReader(new InputStreamReader(powerShellProcess2.getErrorStream()));
        String line = stdInput.readLine();
        System.out.println("Output");
        if(line != null){
            System.out.println(line);
            JOptionPane.showMessageDialog(null,line + System.lineSeparator() + "Please manually email this file to Robert.");
        }
        stdInput.close();
        powerShellProcess2.getOutputStream().close();
        powerShellProcess2.waitFor();
        System.out.println("ps done");

        //close GUI
        form.frame.dispose();
    }
}
