package com.test.common;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

public class LoginCommon {
    public static void login(WebDriver driver, String username, String password) throws InterruptedException {
        driver.findElements(By.name("tj_login")).get(1).click();
        driver.findElement(By.id("TANGRAM__PSP_11__footerULoginBtn")).click();
        driver.findElement(By.name("userName")).sendKeys(username);
        driver.findElement(By.name("password")).sendKeys(password);
        driver.findElement(By.id("TANGRAM__PSP_11__submit")).click();
    }
}
