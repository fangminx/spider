import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;

import java.util.Iterator;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import static org.junit.Assert.fail;

public class a {
    private WebDriver driver;
    private String baseUrl;
    private boolean acceptNextAlert = true;
    private StringBuffer verificationErrors = new StringBuffer();

    @Before
    public void setUp() throws Exception {
        System.setProperty("webdriver.chrome.driver", "C:\\soft\\chromedriver.exe");
        driver = new ChromeDriver();
        baseUrl = "https://www.google.com/";
        driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
        driver.manage().window().maximize();
    }

    @Test
    public void testUntitledTestCase() throws Exception {
        driver.get("https://www.baidu.com/");
        driver.findElement(By.linkText("新闻")).click();

        String currentWindow = driver.getWindowHandle();
        Set<String> handles = driver.getWindowHandles();
        Iterator<String> it = handles.iterator();
        while (it.hasNext()) {
            String handle = it.next();
            if (currentWindow.equals(handle))
                continue;
            driver.switchTo().window(handle);
        }

        // ERROR: Caught exception [ERROR: Unsupported command [selectWindow | win_ser_1 | ]]
        driver.findElement(By.xpath("(//a[contains(text(),'国际')])[2]")).click();
        driver.findElement(By.linkText("美国惨遭盟友背叛，转手将70亿赠与俄罗斯，特朗普麻..")).click();
        // ERROR: Caught exception [ERROR: Unsupported command [selectWindow | win_ser_2 | ]]
        driver.close();
        // ERROR: Caught exception [ERROR: Unsupported command [selectWindow | win_ser_1 | ]]
    }

    @After
    public void tearDown() throws Exception {
        driver.quit();
        String verificationErrorString = verificationErrors.toString();
        if (!"".equals(verificationErrorString)) {
            fail(verificationErrorString);
        }
    }


    private boolean isElementPresent(By by) {
        try {
            driver.findElement(by);
            return true;
        } catch (NoSuchElementException e) {
            return false;
        }
    }

    private boolean isAlertPresent() {
        try {
            driver.switchTo().alert();
            return true;
        } catch (NoAlertPresentException e) {
            return false;
        }
    }

    private String closeAlertAndGetItsText() {
        try {
            Alert alert = driver.switchTo().alert();
            String alertText = alert.getText();
            if (acceptNextAlert) {
                alert.accept();
            } else {
                alert.dismiss();
            }
            return alertText;
        } finally {
            acceptNextAlert = true;
        }
    }
}
