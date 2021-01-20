# Introduction

- This backtesting technical analysis makes use of "Ichimoku Kinko Hyo Equilibrium" (GOC) from list of stocks to screen potential stocks triggering 'BUY' signal.
- Local path for generating output to MS Excel file and list of security codes can be customized. NASDAQ 100 is used in this example.


# Strategy (Ichimoku Kinko Hyo Equilibrium)

- Buy if Conv > Base on T-day; and Conv < Base on T-1 day, where
- Conv = (Highest High over Period1 + Lowest Low over Period1) divided by 2
- Base = (Highest High over Period2 + Lowest Low over Period2) divided by 2
- Period1 = 9 & Period2 = 26    
