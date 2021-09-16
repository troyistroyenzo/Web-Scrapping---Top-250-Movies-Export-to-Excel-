
# -*- coding: utf-8 -*-

"""
Author: Pastoral, Lorenzo Troy
Date Created: 15/09/2021
Description: This is a simple python project that extracts the top 250 movies from 
IMDb.com and exports them to an excel file
"""
__version__ = "1.0.1"
__email__ = "troyenzoo@gmail.com"
__status__ = "Production"
#--------------------------------------#

# Main Modules
from bs4 import BeautifulSoup
import requests
import openpyxl

excel = openpyxl.Workbook()  # Create new excel file
sheet = excel.active
sheet.title = 'Top 250 Movies'
sheet.append(['Movie Rank', 'Movie Name', 'Year of Release', 'Rating'])
print(excel.sheetnames)


# Get Source, check if website responds. Then Scrape values / data from website
try:
    source = requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status()  # Captures Error

    # Return and Parse HTML source code
    soup = BeautifulSoup(source.text, 'html.parser')

    # Look for class and returns source code
    movies = soup.find('tbody', class_="lister-list").find_all('tr')

    # Loop through movies that have the tag 'td' with class "Title Column"
    for movie in movies:
        # Extract HTML from tag
        name = movie.find('td', class_="titleColumn").a.text  # Extract name
        rank = movie.find('td', class_="titleColumn").get_text(
            strip=True).split('.')[0]  # Extract Rank
        year = movie.find('td', class_="titleColumn").span.text.strip(
            '()')  # Extract Year
        # Extract reating
        rating = movie.find('td', class_="ratingColumn imdbRating").strong.text

        print(rank, name, year, rating)  # Print data
        sheet.append([rank, name, year, rating])  # Append data to excel


except Exception as e:
    print(e)


# Import and Save To Excel
excel.save('IMBD Movie Ratings.xlsx')
print("Exported!")
