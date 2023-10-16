# Automated Web Scraper to fetch details of data science based startups in UAE for Expand North Star 2023 and store in Excel for Job Applications

Website - https://exhibitors.expandnorthstar.com/north-star-2023/Exhibitor

Created an automated webscraper using Python and Selenium to extract Company Name, Description, Website and LinkedIn information of all companies in the data science domain in the United Arab Emirates (UAE) for Expand North Star 2023 to use in the future for job applications. First search keyword-'data' and country - 'United Arab Emirates' is used and all relevant startup details are fetched. Then we repeat the same process for the product sector 'Big Data & Analytics'. All the above data is concatenated into a dataframe and duplicate rows are removed. Finally, the dataframe is stored in an excel sheet with 4 columns - 'Company Name', 'Description', 'Website', 'LinkedIn'. 
