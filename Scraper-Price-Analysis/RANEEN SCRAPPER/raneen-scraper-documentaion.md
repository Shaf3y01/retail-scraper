## Raneen Scraper Logic Documentation

### 

### 1\. Imports and Setup

Libraries:

Selenium (web automation)

Pandas (Excel I/O)

OpenPyXL (Excel styling)

Standard libraries: time, os, re, datetime

Chrome Driver:

Configured with options (window size, GPU disabled)

Instantiated globally for the script



### 2\. Input Excel Handling

Reads raneen-target-links.xlsx (header at row 1)

Strips whitespace from column names

Drops rows missing "Category" or "URL"

Creates a list of (Category, URL) tuples for scraping



### 3\. Output Directory

Ensures raneen-products directory exists for saving results



### 4\. Excel Styling Function

Applies:

Header fill color (990000), white bold font

Body font black

Center alignment for all cells

Bottom border for all cells

Auto column width based on content



### 5\. Helper Functions

normalize\_price(text):

Extracts digits from price text, returns as integer or None

extract\_sku(name):

Uses regex to extract SKU-like patterns from product name

normalize\_sku(sku):

Removes dashes, underscores, spaces, converts to lowercase



### 6\. Scraping Logic

##### For each (category, url):

Navigates to the URL

Waits 2 seconds for page load

##### Product Loading

Scrolls to the bottom of the page repeatedly

After each scroll, checks if the number of loaded products increases

Stops after 3 consecutive scrolls with no new products

##### Product Extraction

For each product card (div.product-item-info):

Extracts:

Title (a.product-item-link)

Product URL (from link)

New price (span.price-wrapper span.price)

Old price (span.old-price span.price)

SKU (from title)

Normalized SKU

Handles missing prices gracefully



### 7\. Saving Results

##### If products found:

Saves data to Excel file named raneen\_{category}\_{date}.xlsx in output directory

Applies Excel styling

##### If no products found:

Prints warning



### 8\. Cleanup

Quits the Selenium driver after all categories processed



### 9\. Console Output

Prints progress and status messages throughout (category, products loaded, errors, save location, etc.)

