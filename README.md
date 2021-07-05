# Refactoring in Excel VBA
## **Overview of Project**
The project initially began to assist Tom with analyzing a selected stock’s trading value and price over two years.  It evolved into developing an analytical tool capable of analyzing a large number of stocks efficiently.  
### **Purpose**
The primary purpose of this project was to progress through the development of a single stock analysis, to analyzing multiple stocks (looping through code previously developed, and then finally through revising the code to efficiently use time and computer resources. 
## Results
### Comparison of Stock Performance – 2017 to 2018
The results produced by the Stock Analysis quickly show that the returns for the majority of the stocks analyzed deteriorated, in fact, experiences negative returns based on stock price.  The code’s formatting as Red and Green based on +/- value greatly facilitated identifying negative returns. 

![2017_2018 Results](https://github.com/honoruru/stock-analysis/blob/main/2017_2018%20Results.png)

However, the output does not lend itself to year-over-year comparisons.
When the two years are shown graphically, one can better compare 2017 to 2018.

![Volume and Returns](https://github.com/honoruru/stock-analysis/blob/main/Volume_Returns%20sideXside.png)

Steve's parents believe that if a stock is traded often (high volume), then the price “will accurately reflect the value of the stock.”
For their sake, trading volume and returns were overlaid in the graph below.  From their point of view, the more heavily traded stocks are most accurately valued.  While it’s tempting to draw conclusions on the relationship of volume to stock price accuracy, or even returns, this would seem to be simply an exercise in demonstrating graphing skills in Excel. 

![Volume and Return overlaid](https://github.com/honoruru/stock-analysis/blob/main/Volume_Returns.png)

### Comparison of Analytical Tool Performance after Refactoring
The final code was an impressive improvement over the initial “looped” code.  

![Timer Comparison](https://github.com/honoruru/stock-analysis/blob/main/Timer%20Comparison.png)

As the processing timer demonstrates, the final code's run time was approximately three-hundredths the original time.  When one considers what would once take 100 days, would now take 3 days, the efficiency gain is tremendous.  With the small number of stocks analyzed, the difference wait time is noticeable but not startling.  When considering analyzing the 5,000 or so publicly traded stocks in the U.S., the difference would be waiting for 20 minutes versus 30 seconds.  If you are running the original code, make sure you start it before your smoke break.

Looking at the original and refactored code side-by-side, one can see that the huge increase in processing efficiency was achieved by three actions:
1.	Replacement of Nested Loops with multiple Non-nested Loops
2.	Use in indexed Arrays
3.	Elimination of unnecessary coded. (Shown as highlighted)

Analytical Code
![Code Comparions Analysis](https://github.com/honoruru/stock-analysis/blob/main/Code%20Comparison%20-%20ANALYSIS.png)

Formatting Code
![Code Comparison Formatting](https://github.com/honoruru/stock-analysis/blob/main/Code%20Comparison%20-%20FORMATTING.png)

## Summary
### Advantages (Pros) of Refactoring the Original VBA Script
The compelling advantage of refactoring code is the multifold increase in performance.  In our process, performance improved by almost a factor of a hundred. 

Another advantage is that refactoring exploits existing code. Why re-invent the wheel?  That code also provides a baseline in performance where one can measure whether improvement was worthwhile.

### Disadvantages (Cons) of Refactoring the Original VBA Script
The disadvantage of starting with existing code is one is locked into a previously defined approach.  Refactoring essentially is an attempt to revise the previous code, whether it be the analysis or formatting, in a more efficient way.  There could, in fact, be a more ingenious approach to the coding than the existing code.  But then that would mean starting from scratch. 

One possible disadvantage of eliminating unnecessary code could be that the original code may have been written with the thought of addressing all foreseeable possibilities.  If code is thought to be unnecessary because it appeared to be unused over a certain sample, it might be discovered that it was useful, and necessary when applied to a population or much larger sample. 

