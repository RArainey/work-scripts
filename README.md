# work-scripts

When I was employed as a shipper/receiver by a telecommunications company, I worked at a small location staffed by me, my forklift, and a few inventory systems. The software I used was simple, but inefficient; too much time was wasted waiting for pages to refresh and reports to be served. I decided to write scripts to automate my more repetitive daily tasks. To do that, I needed a language that could handle Internet Explorer and Excel, and found AutoIT 3. Once I had written several scripts in AutoIT 3, I became curious about other languages, and picked Python to create Cleaner.

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
