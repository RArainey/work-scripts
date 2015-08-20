# work-scripts

I was employeed as a shipper/receiver in a warehouse for a telecommunications company. It was a small location staffed by me and my forklift and a few inventory systems. The software we had to use was simple enough, but there was too much time wasted waiting for pages to refresh and reports to be served. It was frustrating having to sit and watch "Please wait..." everyday. I knew enough about software to be dangerous and set out to make some scripts. I googled around for something that could simply handle Internet Explorer and Excel and found the [AutoIT 3](https://www.autoitscript.com/site/autoit/) language. All scripts are .au3, except [Cleaner](/Cleaner.py), because I was curious about Python.

####Scripts:
---
[ECOPS](/ECOPS.au3): accepts an order number and navigates to a submission form to enter the packing slip details.

[InboundOrders](/InboundOrders.au3): retrieves all inbound orders' metadata from the web then cleans and prints via MS Excel.

[OpenOrders](/OpenOrders.au3): retrieves all outbound orders for all plants. Created for supervisor.

[Cleaner](/Cleaner.py): returns serial numbers in .txt from OCR output. The SNs are input for SerialInput.

####Helper Scripts:
---
[YearsFromToday](/YearsFromToday.au3): returns the current date, with adjusted year, in MM/DD/YYYY format.

[CloseHiddenIE](/CloseHiddenIE.au3): closes all instances of Internet Explorer that are invisible.

[SerialInput](/SerialInput.au3): hacky, intended-to-replace, script that enters serial numbers into inventory system.
