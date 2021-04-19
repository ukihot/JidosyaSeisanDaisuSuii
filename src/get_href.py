import requests, bs4

res = requests.get('https://tonari-it.com')
res.raise_for_status()
soup = bs4.BeautifulSoup(res.text, "html.parser")
elems = soup.select('#list h2 a')
for elem in elems:
    print('{} ({})'.format(elem.getText(), elem.get('href')))