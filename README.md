# Web-Scraping-MTG
A project in September – November 2024. I decided to sell most of my Magic: the Gathering collection, and I wanted to figure out where I could get the most value from. I built my collection out in Moxfield and downloaded it as a spreadsheet. Then I identified several websites which have buylists listed online and built a web scraper for each and a system to tell me which card to sell to which site. 

In my case, I wanted store credit from each website. You would have to investigate the HTML to find the cash section if you wanted to be paid out in cash instead. 

The scraper can be inconsistent for ABU Games and Cape Fear Games. The former lists foreign versions of high value cards, which causes wrapping onto an additional page that my scraper does not see. The latter uses fuzzy matching in the buylist search, and the similar results can also push the results you want onto a hidden additional page. Additionally, those websites don’t always display the collector number, which is helpful on cards with multiple treatments. 

Known limitations:

-Cards beyond the first page of searches

-Cards with non-Latin characters (usually from the Lord of the Rings set – change your accented letters to normal letters in the spreadsheet and they work fine)

-Promos and tokens (because Moxfield lists them as a different set than the websites use)

-I’m assuming you’re selling all Near Mint cards or don’t care if one website gives a worse price for near mint but a better price for played.
