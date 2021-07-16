from urllib.request import Request, urlopen

url="https://www.cmegroup.com/markets/energy/crude-oil/light-sweet-crude.quotes.html"
req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})

web_byte = urlopen(req).read()

webpage = web_byte.decode('utf-8')