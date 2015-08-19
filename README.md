# work-scripts

I was employeed to maintain a warehouse for a telecommunications company. It was a small location staffed by me and my forklift and a few inventory systems. The software was simple enough to use, but there was too much waiting for pages to refresh and reports to be served. I knew enough to be dangerous and set out to make some scripts. I googled around for something that could simply handle Internet Explorer and Excel and found the [AutoIT 3](https://www.autoitscript.com/site/autoit/) language. All but Cleaner.py are .au3 scripts.

####Scripts:
---
ECOPS: accepts an order number and navigates to a submission form to enter the packing slip details.

InboundOrders: retrieves order metadata for the selected warehouse from the web then cleans and prints via MS Excel.


####Helper Scripts:
---
[YearsFromToday](/YearsFromToday.au3): Returns the current date, with adjusted year, in MM/DD/YYYY format.

[Cleaner](/Cleaner.py): Returns serial numbers in .txt from OCR output.
