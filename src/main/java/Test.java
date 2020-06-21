import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.util.List;
import java.util.Set;

public class Test {
    private static WebDriver driver;
    public static void main(String[] args) throws InterruptedException {
        //设置系统变量，指定驱动文件路径
        System.setProperty("webdriver.chrome.driver", "C:\\soft\\chromedriver.exe");
        //初始化driver驱动对象
        driver = new ChromeDriver();
        //访问拉钩主页获取获取
        landLagouAndFetchData();
    }

    /**访问拉钩主页获取获取
     * @throws InterruptedException
     */
    private static void landLagouAndFetchData() throws InterruptedException {
        //访问百度
        driver.get("http://www.baidu.com");
        //窗口最大化
        driver.manage().window().maximize();
        //输入查找内容
        driver.findElement(By.id("kw")).sendKeys("拉钩");
        //点击"百度一下"完成搜索
        driver.findElement(By.id("su")).click();
        //获取当前页面的句柄
        String currentHandle = driver.getWindowHandle();
        //点击进入官网
        getElement(By.partialLinkText("互联网求职招聘找工作-上拉勾招聘-专业的互联网求职招聘...")).click();
        //获取或有打开页面的句柄
        Set<String> handles = driver.getWindowHandles();
        //循环句柄
        for (String handle : handles) {
            //获取拉钩页面的句柄
            if(!handle.equals(currentHandle)){
                driver.switchTo().window(handle);
            }
        }
        try {
            //选择全国站
            getElement(By.partialLinkText("全国站")).click();
        } catch (Throwable e) {
            //失败重试机制
            landLagouAndFetchData();
        }
        //输入查询数据
        getElement(By.id("search_input")).sendKeys("自动化测试工程师");
        //点击"搜索"按钮，完成搜索
        getElement(By.id("search_button")).click();

        //因为所有岗位信息显示区域样式都一样，因此根据样式获取每页的所有岗位信息<li>...</li>
        List<WebElement> webElements =getElements(By.className("default_list"));
        //循环处理
        for (WebElement webElement : webElements) {
            //打印岗位信息，此处也可考虑保存到文件，后面再统一写入到文件
            System.out.println(webElement.getText());
            //分隔符
            System.out.println("-------------------------------");
        }
        //循环拿到其余每一页的数据
        while(true){
            //获取"下一页"按钮
            WebElement nextPage = getElement(By.className("pager_next"));
            //循环执行条件："下一页"按钮不为空，且可点击
            if(nextPage!=null&&nextPage.isEnabled()){
                //点击下一页
                nextPage.click();
                //等待5秒
                Thread.sleep(5000);
                //根据样式获取每页的所有岗位信息<li>...</li>
                List<WebElement> webElements2 =getElements(By.className("default_list"));
                //循环打印
                for (WebElement webElement : webElements2) {
                    System.out.println(webElement.getText());
                    System.out.println("-------------------------------");
                }
            }else{
                //不满足条件则终止循环
                break;
            }
        }
    }

    /**显示等待
     * @param locator 元素定位选择器
     * @return
     */
    private static List<WebElement> getElements(By locator) {
        //显示等待wait对象，最长等待时间30秒，默认500毫秒页面轮询一次
        WebDriverWait wait = new WebDriverWait(driver, 30);
        try {
            //调用until方法，传入条件：满足条件的元素列表出现在页面。如果30秒内未找到，则抛出超时异常
            List<WebElement> webElements = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(locator));
            return webElements;
        } catch (Exception e) {
            System.out.println("定位元素超时了");
        }
        return null;
    }

    /**显示等待
     * @param locator 元素定位选择器
     * @return
     */
    public static WebElement getElement(By locator){
        //显示等待wait对象，最长等待时间30秒，默认500毫秒页面轮询一次
        WebDriverWait wait = new WebDriverWait(driver, 30);
        try {
            //调用until方法，传入条件：满足条件的元素出现在页面。如果30秒内未找到，则抛出超时异常
            WebElement webElement = wait.until(ExpectedConditions.presenceOfElementLocated(locator));
            return webElement;
        } catch (Exception e) {
            System.out.println("定位元素超时了");
        }
        return null;

    }
}