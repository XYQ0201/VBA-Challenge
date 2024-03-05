# VBA-Challenge
The initial data set contains opening,closing, daily high and daily low price on each stockmarket trading day for different tickers in 2018,2019 and 2020. 
To summarize the data, I first used for loop to compare the ticker names to identify different tickers and print them out.
Once the next ticker is identified. I write codes to summarize, yearly change, percent change and total stock volumn data.
The yearly change formula is to use the closing price of the last trading day - the openining price of the first day of that year. 
Percent change is equal to (yearly change/opening price) * 100, the column is formatted using FormatPercent to keep 2 decimal points. 
Total Stock Volumn is to add up all the stock volumn on each trading day for the same ticker. 
For bonus points questions, 
I store the first ticker's yearly percent change as initial value for both greatest increase% and greatest decrease%, if next %change is higher than the initial value, this will be deemed as greatest increase, vice versa. I compare the rest of data points line by line to find highest and lowest value.
Same rule applies for greatest total volumn, the initial data is the first total volumn from the summarization, then go through each line to find which one is the biggest. 

Screenshots for results: 
2018:
<img width="1440" alt="2018-1" src="https://github.com/XYQ0201/VBA-Challenge/assets/159677165/2b72ac52-b10c-4ea7-9048-c827cb66217a">
<img width="1440" alt="2018-2" src="https://github.com/XYQ0201/VBA-Challenge/assets/159677165/36597cd3-919e-44e8-8720-488068d99607">
<img width="1440" alt="2018-3" src="https://github.com/XYQ0201/VBA-Challenge/assets/159677165/51c13d5a-21b4-49ef-b55b-d19145e4b323">
2019
<img width="1440" alt="2019-1" src="https://github.com/XYQ0201/VBA-Challenge/assets/159677165/392d0943-c1a0-4286-a578-caea749f52a1">
<img width="1440" alt="2019-2" src="https://github.com/XYQ0201/VBA-Challenge/assets/159677165/5d3237af-8136-4e65-9cb0-27ea3e87692b">
<img width="1440" alt="2019-3" src="https://github.com/XYQ0201/VBA-Challenge/assets/159677165/f0dce518-1c06-452c-b512-f94ae1ed19f5">
2020
<img width="1440" alt="2020-1" src="https://github.com/XYQ0201/VBA-Challenge/assets/159677165/4fa7cca2-bdc0-4a5c-a08e-6b091dea633f">
<img width="1440" alt="2020-2" src="https://github.com/XYQ0201/VBA-Challenge/as<img width="1440" alt="2020-3" src="https://github.com/XYQ0201/VBA-Challenge/assets/159677165/1e0b6ed3-4666-4694-a059-863906d6c571">
sets/159677165/fbc285cd-d1d9-47ef-8d0d-51ed9cae1848">
<img width="1440" alt="2020-3" src="https://github.com/XYQ0201/VBA-Challenge/assets/159677165/05e018ef-76cc-4307-9526-2c9c97f7333c">
