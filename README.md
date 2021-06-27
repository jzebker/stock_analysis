# Stock Analysis

## Overview of Project
Refactor VBA code to help Tom analyze stocks for his parents ***more efficiently***.  In addition, provide a written analysis that details this process and evaluates results.

## Results

### Analysis of Stock Performance in 2017 vs 2018
![returns2017vs2018](https://user-images.githubusercontent.com/84994321/123528398-8d126500-d69b-11eb-927d-442b7a097e7b.png)
Comparing returns from 2017 and 2018, every tracked stock (other than TERP and RUN) either gained less or lost money in 2018 when compared to 2017. TERP lost less money in 2018 than in 2017 and RUN grew 80%. ENPH and RUN were the only tracked stocks with positive returns for 2018.
![tdv2017vs2018](https://user-images.githubusercontent.com/84994321/123528516-9223e400-d69c-11eb-8f15-26de5249c7b3.png)
The following stocks experienced an increase in trade volume from 2017 to 2018: DQ, ENPH, HASI, RUN, SEDG, TERP, and VSLR.  It is worth noting that DQ, ENPH, and RUN all experienced significant proportional increases to their trade volume from 2017 to 2018.  ENPH and RUN were the only stocks that saw positive returns in 2018.  DQ does not fit this pattern but its total volume numbers are relatively small when compared to those of ENPH and RUN.  Data for the charts above follows.
<table class="tg" align="center">
<thead>
  <tr>
    <th class="tg-0pky">2017</th>
    <th class="tg-0pky">2018</th>
  </tr>
</thead>
<tbody>
  <tr>
    <td class="tg-0pky"><img width="240" alt="Screen Shot 2021-06-26 at 4 09 39 PM" src="https://user-images.githubusercontent.com/84994321/123528003-f5ac1280-d698-11eb-8b7b-12f7caffed92.png"></td>
    <td class="tg-0pky"><img width="243" alt="Screen Shot 2021-06-26 at 4 10 43 PM" src="https://user-images.githubusercontent.com/84994321/123528013-15dbd180-d699-11eb-98da-680ec9742914.png"></td>
  </tr>
</tbody>
</table>

### Analysis of VBA Refactoring
Initially, our code took > .6 seconds to run. It utilized a nested loop to loop once through *all of the data* for *each* ticker value in the 'tickers' array.  In this case, it searches through the data 12 times:
```
...
For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        Worksheets(yearValue).Activate
        For j = 2 To rowEnd
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
...
```
Here are run times before refactoring:

<img width="686" alt="oldcode2017" src="https://user-images.githubusercontent.com/84994321/123558847-54cb5f00-d74d-11eb-950f-cf2314329c10.png">

<img width="679" alt="oldcode2018" src="https://user-images.githubusercontent.com/84994321/123558851-59901300-d74d-11eb-8948-1073606f8187.png">

This works fine but we want to scale up and go real fast.
## Summary

### What are the advantages or disadvantages of refactoring code?

### How do these pros and cons apply to refactoring the original VBA script?
