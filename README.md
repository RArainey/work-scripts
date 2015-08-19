# work-scripts
Scripts created to remove tedium from a warehouse management job.

I was employeed to maintain a warehouse for a telecommunications company. It was a small location staffed by me and my forklift and a few inventory systems. The software was simple enough to use, but there was too much waiting for pages to refresh and reports to be served. I knew enough to be dangerous and set out to make these scripts to handle data manipulation and web navigation. I googled around for a language that could simply handle Internet Explorer and Excel and found AutoIT 3.

In order of tedium spared:
ECOPS: accepts an order number and navigates to a submission form to enter the packing slip details.
InboundOrders: retrieves order metadata from web, cleans and prints table via MS Excel.


Helper Scripts:
YearsFromToday: Accepts an int and returns the current date, with adjusted year, in MM/DD/YYYY format.
