# RPA1 – Project Documentation
# Janko Zivaljevic
# 1. Development Environment:
IntelliJ IDEA 2022.2.4 (Community Edition)
Build #IC-222.4459.24, built on November 22, 2022
Runtime version: 17.0.5+7-b469.71 amd64
VM: OpenJDK 64-Bit Server VM by JetBrains s.r.o.
Windows 10 10.0
GC: G1 Young Generation, G1 Old Generation
Memory: 1012M
Cores: 4
# 2. Libraries and Dependencies:
I have decided to use Maven on this project. Reason is that I found Maven really helpful when 
working on a project where external libraries are needed (Selenium, Apache POI, etc…) as in 
this one. Also, I think Maven is extremly helpful in building the right project structure and 
getting right JAR files for each project. Also, at any point of the project we can migrate to 
different version of the dependency,etc… I have downloaded all the dependencies from the 
Remote Maven Repository(https://mvnrepository.com/) and here I am going to explain all of 
them.

# 3. Configuration Instructions:
Inside RunnableJar folder, you will se a configuration file, and inside that file two variables: 
ratingToCompare and numOfMoviesToExtract. You will also see that these two variables are 
already initialized to 9.0 and 20.
By changing numOfMoviesToExtract variable, you will define exactly how many movies you 
want in your Excel file. By default, it is 20 but you can go as far as 250 because that’s how 
many movies are on the list. Changing this variable will affect all the excel files that will be 
created in this project, and it will also affect the value of average rating of all the movies.
By changing ratingToCompare variable, you will change two excel files created in archive 
directory: SubsetOfRatingAbove and SubsetOfRatingUnder as by default rating to compare is 
9.0 and that’s how movies are sorted in these two excel files. For example, if you set this 
variable to 8.5, than you will get all the movies that have rating above or equal to 8.5 in excel 
file SubsetOfRatingAbove, and all the movies under that rating in the excel file 
SubsetOfRatingUnder. 
# 4. Setup Guide:
In order to run this project in IntelliJ IDE, you need to make sure that you have Maven
installed. In order to install it visit the website (https://maven.apache.org/install.html). After 
installation you need to verify that Maven is installed properly by running command mvnversion in terminal. After that you just need to go to File – Open and then select the project 
folder. After that everything will be opened for you in IntelliJ. All the dependencies will be 
imported automatically.
In order to run this project in Eclipse IDE, you just need to click File – Import – Maven – Existing 
Maven Projects and then just locate the project folder. Then click on finish and that will 
sucessfully set up this project for you in Eclipse. If you, for any reason, want to run project 
from IDE and not from jar file, You just need to realocate configuration file and put it outside 
of src folder.
# 5. Execution Guide:
In order to run executable .jar file, all you need to do is unzip the RunnableJAR folder, then 
open “jar file ovde” folder, and inside that folder you will see Executable Jar File named 
breznaProjekatTest-1.0-SNAPSHOT-jar-with-dependencies, which then you just need to click 
and project will start executing. Also, in the same folder you can see configuration file which 
you can modify by changing Rating to Compare and Number of Movies that you want to 
extract. Than inside src-main-resources, your RPA Task folder will be created with directories 
Archive and Working, and inside Archive directory you will be able to find all the excel subsets.
# 6. Code Structure and Functionallity:
Once you run .jar file, you will be sent to the IMDB webpage where I scraped Top 250 Movies 
list, after that you will be spent to RPA Task website where in the input field, average rating 
of the movies will show up. Also, inside main folder, RPA Task folder will be created and all 
other needed Excel files.
I have created two classes in my project, and used built-in Maven structure. Class XLUtility is 
a helper class that just makes my life easier. Whenever I need to manipulate with .xlsx files, I 
just call object of the XLUtility class. Class Main is where I have developed all the project, it is 
separated into logical methods and I tried to make my code as clean as possible. Also, I tried 
to make my code easy to extend and to maintain. I have used comments in almost all the 
methods and I hope there will be no problems in understanding the logic behind how I did 
this task
