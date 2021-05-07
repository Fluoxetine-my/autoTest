package com.test.cases;

import com.test.common.LoginCommon;
import com.test.utils.ExcelReport;
import com.test.utils.ReadExcel;
import com.test.utils.SetUP;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.Reporter;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

import static com.test.utils.SetUP.getBaseUrl;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.testng.Assert.*;

public class LoginCase {
    private WebDriver driver;
    private StringBuffer verificationErrors = new StringBuffer();
    private String URL= getBaseUrl();
    //获取包名
    private String packageName=this.getClass().getPackage().getName();
    //获取类名
    private String className=this.getClass().getName();
    //测试前初始化
    @BeforeClass(alwaysRun = true)
    public void setUp() throws Exception {
        SetUP login = new SetUP();
        login.setProperty();
        //设置浏览器属性
        ChromeOptions options = SetUP.setChromeOption();
        //初始化driver
        driver = new ChromeDriver(options);
    }

    //用数组接收从TestData中读取的数据,这里要注意，一个xlsx文件可能有很多sheet表，所以sheetName要对应
    @DataProvider(name = "Login")
    public Object[][] Login() throws IOException {
        return ReadExcel.getData("src\\test\\java\\com\\test\\datas","TestData.xlsx","login");
    }

    //测试用例，传入参数是从DataProvider中遍历获取的数据
    @Test(dataProvider = "Login")
    public void LoginTest(String name,String password) throws Exception {
        Reporter.log("测试用例：登录");
        driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
        //登录
        System.out.println("0.登录百度");
        //第1步：获取路径，进入登录页面,判断是否进入
        driver.get(getBaseUrl());
        System.out.println("进入百度登录界面：" + getBaseUrl().equals(driver.getCurrentUrl()));
        //获取当前方法名
        String methodName=Thread.currentThread().getStackTrace()[1].getMethodName();

        //第2步：输入正确的用户名和密码，点击登录(调用封装的登录方法)
        LoginCommon.login(driver, name,password);
        //点击登录有个转图片的验证码，这里没有更好的办法，只能暂时延迟手动转
        Thread.sleep(10000);

        //第3步：获取标签，看是否登录成功并写入excel结果
        String message=driver.findElement(By.xpath("//*[@id=\"u1\"]/a")).getText();
        //如果该位置依旧为登录，则表示登录失败
        if(!message.equals("登录")){
            ExcelReport.writeExcel(packageName+"登录测试",className,methodName,"登录","pass","");
        }else{
            ExcelReport.writeExcel(packageName+"登录测试",className,methodName,"登录","fail","登录失败");
        }
    }

    //测试结束关闭driver，收尾
    @AfterClass(alwaysRun = true)
    public void tearDown() throws Exception{
        driver.quit();
        String verificationErrorString = verificationErrors.toString();
        if (!"".equals(verificationErrorString)) {
            fail(verificationErrorString);
        }
    }
}
