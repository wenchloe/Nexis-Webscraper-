# Nexis-Webscraper-

# Description: NexisWebscraper takes in a company name/search query and collects data from every article within the specified parameters from the Nexis Uni database through webscraping. The client can specify a time frame, a search query, and a list of publication types. The program extracts the title, date, publisher, word count, and all of each article's text (every article fitting the search query and the filters). Inputs the data into a given excel workbook and sheet (prints after every page, or after every ten articles collected).

# User-Specified Inputs / Parameters: Time Frame (year and month for cut-off), Search Query, Login Net-Id and Password, Publication Types, Excel Workbook / Sheet

# Pre: Download Apache POI zip files, Commons Collection, geckodriver, and Selenium WebDriver
# 		 Configure build path in Eclipse - add external jars: 
#  		    - Apache POI: poi-3.17.jar
#         - Commons Collection: commons-codec-.1.10.jar, commons-collections4-4.1-javadoc.jar,
#           Commons-collections4-4.1.jsar,commons-exec-1.3.jar, commons-logging-1.2.jar
#         - Selenium WebDriver: selenium standalone server, client-combined-3.8.1.jar, client-combined-3.8.1-sources.jar
