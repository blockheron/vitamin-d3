This project was my introduction to web scraping as a part of my one-month internship at Kalbe Farma. It's an automated program that will search up products with "vitamin d3" in their names in the BPOM database, scrapes data, then saves the data to an Excel file. The program name is "scraping_test.py", while the resulting Excel file is "produk_vitD3.xlsx". Please note that its contents are largely in Indonesian.

The pipeline for this web scraping process is:
1. Navigate to "vitamin d3" search results.
2. Scrape information from each search result, iterating through all the pages while doing so. After scraping one page (that contains 10 results), write it to an Excel file.
4. One all pages have been scraped, "click" each search result, which pops open new information, and scrape information from there one-by-one. In actuality, the "click" is a Javascript injection.
5. Write to an Excel file for every 5 clicked/injeted results.

If the program quits while going thorugh clicked/injected results due to a timeout error (or any other error), it can continue searching from where it left off upon re-running the program. It'll read the code (labelled "kode" in the program) of the last written item, find its index position in the first main sheet containing the page results, get the code of the next index position from the page results, and use that to continue sarching through clicked/injected results.

While it fully works and is finished, it can still improved:
1. If a search result is categorised as something other than "SK" (Suplemen) or "KO" (Kosmetika), the program will fail. Each category stores its data in a different table, and the program does not accommodate categories other than "SK and "KO".
2. One of the last search results in the "SK" category has a column that no other items in that category had. Due to time constraints, I didn't address this bug.
3. There is no recovery in the first half of the program where the program iterates through search result pages. I made this decision because this part ran faster and had significantly less timeout errors, and also because of time constraints. However, having no recovery for a search result that contains significantly more pages is risky.
4. Other small optimisations in the code can be made. For the purposes of learning web scraping and, again, due to time constraints, I didn't address this.

I will use this program in a new project that builds on top of what I've made and learned so far. To prepare for this, I plan to write separate functions for each search result category in its own file, so that the program can more easily accommodate more categories and to increase the code's readibility. 
