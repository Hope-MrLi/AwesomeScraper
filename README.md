# Awesome Scraper
> A handy tool for scraping Tianyancha. 

![image](https://github.com/mtyylx/AwesomeScraper/blob/master/snapshot.png?raw=true)

## Dependency

1. PhantomJS / Chromedriver

2. Selenium

3. BeautifulSoup

4. TKinter


## Implementation Details

1. Reading from a user specified txt file as input (UTF-8).

2. Generate corresponding search link.

3. Browse each link on Tianyancha using Selenium and PhantomJS/Chromedriver.

4. Parse the search page and decide if there are valid return result. Open the first result link to enter company info page.

5. Use BeautifulSoup to parse the html page and extract POIs and store them.

6. Display the current parsed result on the GUI.

7. When all entry in the user specified txt has been searched, return a file that save all the company info.

8. Providing a one-click function for user to post-processing the result file generated, and return a .csv file for user to further analyze in Excel.

9. User can abort and restart the scraper at any time. There are mainly three threads running. Main thread is the GUI which is based on TKinter, and when user start the search, thread Scraper and thread Monitor will be created then destroyed.