import io.github.bonigarcia.wdm.WebDriverManager;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.*;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

public class Main {

    public static String extractValue(String input) {
        int startIndex = input.indexOf('(');
        int endIndex = input.indexOf(')');
        if (startIndex != -1 && endIndex != -1 && startIndex < endIndex) {
            return input.substring(0, startIndex);
        }
        return null;
    }
    static XLUtility xlutilListOfTopMovies;
    private XLUtility xlutilSubsetOfRatingAbove;
    private XLUtility xlutilSubsetOfRatingUnder;
    private XLUtility xlutilSubsetOfMoviesWithThe;
    private XLUtility xlutilSubsetOfMoviesWithoutThe;
    static WebDriver driver = new ChromeDriver();

    public Main(){
        xlutilListOfTopMovies = new XLUtility(WORKING_PATH+"IMDB.xlsx");
        xlutilSubsetOfRatingAbove = new XLUtility(ARCHIVE_PATH+"SubsetOfRatingAbove.xlsx");
        xlutilSubsetOfRatingUnder = new XLUtility(ARCHIVE_PATH+"SubsetOfRatingUnder.xlsx");
        xlutilSubsetOfMoviesWithThe = new XLUtility(ARCHIVE_PATH+"SubsetOfMoviesWithThe.xlsx");
        xlutilSubsetOfMoviesWithoutThe = new XLUtility(ARCHIVE_PATH+"SubsetOfMoviesWithoutThe.xlsx");
    }

    private static final String BASE_PATH = "src/main/resources/RPA Task/";
    private static final String WORKING_PATH = BASE_PATH+"working/";
    private static final String ARCHIVE_PATH = BASE_PATH+"archive/";

    public String readConfigFile(String propertyName) throws IOException {
        try(FileReader reader = new FileReader("configuration")) {
            Properties properties = new Properties();
            properties.load(reader);
            return properties.getProperty(propertyName);
        }
    }

    public ArrayList<Double> convertToDouble(ArrayList<String> ls){
        ArrayList<Double> doublelist = new ArrayList<>();
        for(String stringVal:ls){
            try{
                double doubleVal = Double.parseDouble(stringVal);
                doublelist.add(doubleVal);

            }catch(NumberFormatException e){
                e.printStackTrace();
            }
        }
        return doublelist;
    }

    public ArrayList<String> listOfElementsInCertainCol(int colNum) throws IOException {
        String numOfMoviesToExtract = readConfigFile("numOfMoviesToExtract");
        ArrayList<String> ls = new ArrayList<String>();
        for(int i=1;i<=Double.parseDouble(numOfMoviesToExtract);i++) {
            String myData = null;
            myData = xlutilListOfTopMovies.getCellData("IMDB Movies", i, colNum);
            ls.add(myData);
        }
        return ls;
    }

    public void createDir() throws FileNotFoundException {
        File file = new File("src/main/resources/RPA Task");
        file.mkdirs();
        File file2 = new File("src/main/resources/RPA Task/working");
        file2.mkdirs();
        File file3 = new File("src/main/resources/RPA Task/archive");
        file3.mkdirs();
    }

    public void setExcelHeader(XLUtility xlutil, String sheetName, String[] headerLabels) throws IOException{
        for(int i=0;i<headerLabels.length;i++){
            xlutil.setCellData(sheetName,0,i,headerLabels[i]);
        }
    }

    public void createExcelHeader() throws IOException {
//        In this function I am creating header for all the Excel files that I have in this project
//        I am calling setExcelHeader method that is using XLUtil method setCellData
        String[] imdbHeader = {"Ordinal Num","Movie Title","Movie Rating","IMDB Link"};
        setExcelHeader(xlutilListOfTopMovies,"IMDB Movies",imdbHeader);

        String[] sheet1Header = {"Ordinal Num","Movie Title","Movie Rating above set condition","IMDB Link"};
        setExcelHeader(xlutilSubsetOfRatingAbove,"Sheet 1",sheet1Header);

        String[] sheet2Header = {"Ordinal Num","Movie Title","Movie Rating under set condition","IMDB Link"};
        setExcelHeader(xlutilSubsetOfRatingUnder,"Sheet 1",sheet2Header);

        String[] sheet3Header = {"Ordinal Num","Movie Titles with set condition","Movie Rating","IMDB Link"};
        setExcelHeader(xlutilSubsetOfMoviesWithThe,"Sheet 1",sheet3Header);

        String[] sheet4Header = {"Ordinal Num","Movie Titles without set condition","Movie Rating","IMDB Link"};
        setExcelHeader(xlutilSubsetOfMoviesWithoutThe,"Sheet 1",sheet4Header);
    }

    public void movieRating() throws IOException {
//        This function calculates the average of all the movies, and then
//        locates the element on Brezna webpage and submits the result
        WebDriver driver = initializeDriver();
        driver.get("https://docs.google.com/forms/d/1uSkQclzgYimoODTeHpvt2MO7QTAEanQahYGDmqsCkFs/edit");
//        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
        double result = listOfElementsInCertainCol(2).stream()
                .mapToDouble( s -> Double.parseDouble( s ) ).sum();
//        Here I am reading number of movies that I need to extract from the config file
        int numOfMovies = Integer.parseInt(readConfigFile("numOfMoviesToExtract"));
        double avgOfRatings = (double)result/numOfMovies;
        DecimalFormat df = new DecimalFormat("#.##");
        driver.findElement(By.xpath("//input[@type='text']")).sendKeys(df.format(avgOfRatings));
        driver.findElement(By.xpath("//span[contains(text(),'Проследи')]")).click();
    }
    public void splitImportedFileBasedOnRating() throws IOException {
//        Inside this function, I am putting data in two excel subsets: above and under rating to compare
//        Its worth mentioning that you can modify rating to compare via config file
        ArrayList<String> listOfOrdinals = listOfElementsInCertainCol(0);
        ArrayList<String> listOfMovieTitles = listOfElementsInCertainCol(1);
        ArrayList<Double> listOfAvgRating = convertToDouble(listOfElementsInCertainCol(2));
        ArrayList<String> listOfMovieLinks = listOfElementsInCertainCol(3);
        int rowNum1=1;
        int rowNum2=1;

        String ratingToCompare = readConfigFile("ratingToCompare");
        for(int i=0;i<listOfAvgRating.size();i++){
            String ordinals = String.valueOf(listOfOrdinals.get(i));
            String avgValue = String.valueOf(listOfAvgRating.get(i));
            String movieTitle = listOfMovieTitles.get(i);
            String movieLink = listOfMovieLinks.get(i);
            if(listOfAvgRating.get(i) >= Double.parseDouble(ratingToCompare)){
                xlutilSubsetOfRatingAbove.setCellData("Sheet 1",rowNum1,0,ordinals);
                xlutilSubsetOfRatingAbove.setCellData("Sheet 1",rowNum1,1,movieTitle);
                xlutilSubsetOfRatingAbove.setCellData("Sheet 1",rowNum1,2,avgValue);
                xlutilSubsetOfRatingAbove.setCellData("Sheet 1",rowNum1,3,movieLink);
                rowNum1++;
            }
            else {
                xlutilSubsetOfRatingUnder.setCellData("Sheet 1",rowNum2, 0,  ordinals);
                xlutilSubsetOfRatingUnder.setCellData("Sheet 1",rowNum2, 1, movieTitle);
                xlutilSubsetOfRatingUnder.setCellData("Sheet 1",rowNum2, 2, avgValue);
                xlutilSubsetOfRatingUnder.setCellData("Sheet 1",rowNum2, 3, movieLink);
                rowNum2++;
            }
        }
        rowNum1=1;
        rowNum2=1;

    }

    public void splitImportedFileBasedOnAlphabet() throws IOException {
//        Here I am creating two more subsets
//        based on whether movie has word "the" inside the heading
        ArrayList<String> listOfOrdinals = listOfElementsInCertainCol(0);
        ArrayList<String> listOfMovieTitles = listOfElementsInCertainCol(1);
        ArrayList<Double> listOfAvgRating = convertToDouble(listOfElementsInCertainCol(2));
        ArrayList<String> listOfMovieLinks = listOfElementsInCertainCol(3);
        int rowNum1=1;
        int rowNum2=1;
        for(int i=0;i<listOfMovieTitles.size();i++) {
            String ordinals = String.valueOf(listOfOrdinals.get(i));
            String avgValue = String.valueOf(listOfAvgRating.get(i));
            String movieTitle = listOfMovieTitles.get(i);
            String movieLink = listOfMovieLinks.get(i);
            if (movieTitle.toLowerCase().contains("the")) {
                xlutilSubsetOfMoviesWithThe.setCellData("Sheet 1", rowNum1, 0, ordinals);
                xlutilSubsetOfMoviesWithThe.setCellData("Sheet 1", rowNum1, 1, movieTitle);
                xlutilSubsetOfMoviesWithThe.setCellData("Sheet 1", rowNum1, 2, avgValue);
                xlutilSubsetOfMoviesWithThe.setCellData("Sheet 1", rowNum1, 3, movieLink);
                rowNum1++;
            } else {
                xlutilSubsetOfMoviesWithoutThe.setCellData("Sheet 1", rowNum2, 0, ordinals);
                xlutilSubsetOfMoviesWithoutThe.setCellData("Sheet 1", rowNum2, 1, movieTitle);
                xlutilSubsetOfMoviesWithoutThe.setCellData("Sheet 1", rowNum2, 2, avgValue);
                xlutilSubsetOfMoviesWithoutThe.setCellData("Sheet 1", rowNum2, 3, movieLink);
                rowNum2++;
            }
        }
    }
    public static WebDriver initializeDriver(){
        WebDriverManager.chromedriver().setup();
        driver.manage().timeouts().implicitlyWait(3, TimeUnit.SECONDS);
        driver.manage().window().maximize();
        return driver;
    }
    public static void extractMovieData(WebDriver driver, Main main) throws IOException{
//        This is function that navigates to IMDB Top 250 Movies website and then scrapes the data to excel file
        WebDriverManager.chromedriver().setup();
        driver.manage().timeouts().implicitlyWait(3, TimeUnit.SECONDS);
        driver.manage().window().maximize();
        driver.get("https://www.imdb.com/");
        driver.findElement(By.xpath("//label[@id='imdbHeader-navDrawerOpen']")).click();
        driver.findElement(By.linkText("Top 250 Movies")).click();
        String numOfMoviesToExtract = main.readConfigFile("numOfMoviesToExtract");
        for (int r = 1; r <= Double.parseDouble(numOfMoviesToExtract) ;r++) {
            String movieTitle = driver.findElement(By.cssSelector("body > div:nth-child(2) > main:nth-child(6) > div:nth-child(1) > div:nth-child(5) > section:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > ul:nth-child(2) > li:nth-child("+r+") > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > a:nth-child(1) > h3:nth-child(1)")).getText();
            String[] parts = movieTitle.split("\\.");
            String ordinalNum="";
            String actualTitle="";
            if(parts.length >= 2) {
                ordinalNum = parts[0].trim();
                actualTitle = parts[1].trim();
            }
            String movieRating = driver.findElement(By.cssSelector("body > div:nth-child(2) > main:nth-child(6) > div:nth-child(1) > div:nth-child(5) > section:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > ul:nth-child(2) > li:nth-child("+r+") span[class='ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating']")).getText();
            String actualRating = extractValue(movieRating);

            String movieLink = String.valueOf(driver.findElement(By.cssSelector("body > div:nth-child(2) > main:nth-child(6) > div:nth-child(1) > div:nth-child(5) > section:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > ul:nth-child(2) > li:nth-child("+r+") a[class='ipc-title-link-wrapper']")).getAttribute("href"));
            xlutilListOfTopMovies.setCellData("IMDB Movies", r, 0, ordinalNum);
            xlutilListOfTopMovies.setCellData("IMDB Movies", r, 1, actualTitle);
            xlutilListOfTopMovies.setCellData("IMDB Movies", r, 2, actualRating);
            xlutilListOfTopMovies.setCellData("IMDB Movies", r, 3, movieLink);
        }
    }
    public static void main(String[] args) throws IOException {
        Main m = new Main();
        m.createDir();
        m.createExcelHeader();
        WebDriver driver = initializeDriver();
        extractMovieData(driver,m);
        m.movieRating();
        driver.close();
        m.splitImportedFileBasedOnAlphabet();
        m.splitImportedFileBasedOnRating();
        xlutilListOfTopMovies.deleteSheet();
    }
}
