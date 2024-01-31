from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top Rated Movies'
sheet.append(['Movie Name', 'Year of Release', 'IMDB Rating'])

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

base_url = 'https://www.imdb.com/chart/top/'
genres = ['drama', 'adventure', 'thriller', 'action', 'crime', 'comedy', 'mystery', 'war', 'fantasy', 'sci-fi']

print('*******WELCOME TO IMDB TOP MOVIES GENERATOR*******')

while(True): 
    print('\n You can choose among the following genres: ')
    print('1. All 2. Drama 3. Adventure 4. Thriller 5. Action 6. Crime 7. Comedy 8. Mystery 9. War 10. Fantasy 11. Sci-Fi ')
    val = input('\n Type your genre : ').lower()
    url = ''

    while(True):
        if val in genres: 
            url = base_url + '?genres=' + val
            break
        elif val == 'all': 
            url = base_url
            break
        else :
            val = input('\n Oops! Try Again : ').lower()

    try : 
        source = requests.get(url, headers=headers)
        source.raise_for_status()

        soup = BeautifulSoup(source.text, 'html.parser')
        
        movies = soup.find('ul', class_='ipc-metadata-list ipc-metadata-list--dividers-between sc-71ed9118-0 kxsUNk compact-list-view ipc-metadata-list--base').find_all('li')
        # print(len(movies))

        for movie in movies: 
            name = movie.find('h3', class_='ipc-title__text').text
            year = movie.find('span', class_='sc-1e00898e-8 hsHAHC cli-title-metadata-item').text
            rating = movie.find('span', class_='ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating').text
            print(name + ' ' + year + ' ' + rating)
            sheet.append([name, year, rating])

        again = input('Do you wanna try again [y/n?]')
        if again.lower() == 'n': 
            print('\nThank you')
            break

    except Exception as e: 
        print(e)

excel.save('IMDB Movie Ratings.xlsx')

