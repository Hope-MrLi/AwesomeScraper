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

6. Display the current parsed result on the GUI, and record search result in output files. (Producer-Consumer Model)

7. Providing a one-click function for post-processing / duplicate removal, and return *.csv file for user to further analyze in Excel.

8. User can abort and restart the scraper at any time. Three threads: 
    - Main thread: GUI based on TKinter.
    - Scraper thread: Perform search and extraction. 
    - Monitor thread: Keep track of Scraper thread and display most recent search results to GUI.
    
## Code Freeze

```
python setup.py py2exe
```

## Have Fun!

