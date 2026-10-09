"""Builds finance-calculator.csv: a CSV full of live formulas for Sheets."""
import csv
import os

cells = {}


def put(addr, value):
    cells[addr] = value


def col(c):
    return "ABCDEFGHIJ"[c]


# ---------------------------------------------------------------- header
put("A1", "PERSONAL FINANCE CALCULATOR")
put("D1", '=TEXT(TODAY(),"dddd, mmmm d, yyyy")')
put("A2", "Change any INPUT (column B) and everything recalculates.")

# ---------------------------------------------------------------- mortgage
put("A4", "MORTGAGE")
for r, (label, val) in enumerate([
    ("INPUT Home price", "$450,000"),
    ("INPUT Down payment", "20%"),
    ("INPUT Interest rate (APR)", "6.5%"),
    ("INPUT Loan term (years)", "30"),
    ("INPUT Property tax / year", "$5,400"),
    ("INPUT Insurance / year", "$1,800"),
    ("Loan amount", "=B5*(1-B6)"),
    ("Monthly principal + interest", "=ROUND(-PMT(B7/12,B8*12,B11),2)"),
    ("Monthly total incl. tax + ins.", "=ROUND(B12+(B9+B10)/12,2)"),
    ("Total paid over the loan", "=ROUND(B12*B8*12,2)"),
    ("Total interest", "=ROUND(B14-B11,2)"),
    ("Interest as % of loan", '=TEXT(B15/B11,"0.0%")'),
    ("Paid off in", '=TEXT(EDATE(TODAY(),B8*12),"mmmm yyyy")'),
], start=5):
    put(f"A{r}", label)
    put(f"B{r}", val)

put("D4", "FIRST YEAR OF PAYMENTS")
for i, h in enumerate(["Month", "Payment", "Interest", "Principal", "Balance"]):
    put(f"{col(3 + i)}5", h)
for r in range(6, 18):
    put(f"D{r}", "1" if r == 6 else f"=D{r - 1}+1")
    put(f"E{r}", "=$B$12")
    prev_balance = "$B$11" if r == 6 else f"H{r - 1}"
    put(f"F{r}", f"=ROUND({prev_balance}*$B$7/12,2)")
    put(f"G{r}", f"=E{r}-F{r}")
    put(f"H{r}", f"=ROUND({prev_balance}-G{r},2)")
put("D18", "Year 1 totals")
for c in "EFG":
    put(f"{c}18", f"=SUM({c}6:{c}17)")
put("H18", '="Paid off: "&TEXT(1-H17/B11,"0.0%")')

# ---------------------------------------------------------------- savings
put("A20", "SAVINGS & INVESTING")
for r, (label, val) in enumerate([
    ("INPUT Starting balance", "$10,000"),
    ("INPUT Monthly contribution", "$750"),
    ("INPUT Expected annual return", "7%"),
    ("INPUT Years", "25"),
    ("INPUT Inflation", "3%"),
    ("Future value", "=ROUND(FV(B23/12,B24*12,-B22,-B21),2)"),
    ("Total you put in", "=B21+B22*12*B24"),
    ("Growth from returns", "=ROUND(B26-B27,2)"),
    ("Future value in today's money", "=ROUND(B26/(1+B25)^B24,2)"),
    ("Monthly income (4% rule)", "=ROUND(B26*0.04/12,2)"),
    ("Years to double (rule of 72)", "=ROUND(72/(B23*100),1)"),
], start=21):
    put(f"A{r}", label)
    put(f"B{r}", val)

put("D20", "GROWTH BY YEAR")
for i, h in enumerate(["Year", "Start", "Added", "Growth", "End"]):
    put(f"{col(3 + i)}21", h)
for r in range(22, 32):
    put(f"D{r}", "1" if r == 22 else f"=D{r - 1}+1")
    put(f"E{r}", "=$B$21" if r == 22 else f"=H{r - 1}")
    put(f"F{r}", "=$B$22*12")
    put(f"G{r}", f"=ROUND(H{r}-E{r}-F{r},2)")
    put(f"H{r}", f"=ROUND(FV($B$23/12,12,-$B$22,-E{r}),2)")

# ---------------------------------------------------------------- budget
put("A34", "MONTHLY BUDGET")
put("A35", "INPUT Monthly take-home pay")
put("B35", "$7,800")
for i, h in enumerate(["Category", "Budget", "Spent", "Left", "% of pay", "Status"]):
    put(f"{col(i)}36", h)
budget = [
    ("Housing (from mortgage)", "=B13", "=B13"),
    ("Groceries", "650", "712"),
    ("Transport", "400", "365"),
    ("Utilities", "280", "301"),
    ("Dining out", "300", "420"),
    ("Savings (from above)", "=B22", "=B22"),
    ("Health", "200", "150"),
    ("Fun money", "250", "180"),
]
for r, (cat, b, s) in enumerate(budget, start=37):
    put(f"A{r}", cat)
    put(f"B{r}", b)
    put(f"C{r}", s)
    put(f"D{r}", f"=B{r}-C{r}")
    put(f"E{r}", f'=TEXT(C{r}/$B$35,"0.0%")')
    put(f"F{r}", f'=IF(C{r}>B{r},"OVER by "&TEXT(C{r}-B{r},"$#,##0"),"ok")')
put("A45", "Total")
put("B45", "=SUM(B37:B44)")
put("C45", "=SUM(C37:C44)")
put("D45", "=B45-C45")
put("E45", '=TEXT(C45/B35,"0.0%")')
put("F45", '=IF(C45>B35,"Spending more than you earn!","Left over: "&TEXT(B35-C45,"$#,##0"))')
put("A46", "Biggest expense")
put("B46", "=INDEX(A37:A44,MATCH(MAX(C37:C44),C37:C44,0))")
put("A47", "Categories over budget")
put("B47", '=COUNTIF(D37:D44,"<0")')
put("A48", "Average spend per category")
put("B48", "=ROUND(AVERAGE(C37:C44),2)")
put("A49", "Savings rate")
put("B49", '=TEXT(C42/B35,"0.0%")')
put("A50", "Spend excluding housing + savings")
put("B50", "=SUMIFS(C37:C44,A37:A44,\"<>Housing*\",A37:A44,\"<>Savings*\")")

# ---------------------------------------------------------------- quick math
put("H34", "QUICK MATH")
put("H35", "INPUT Any number")
put("I35", "144")
for r, (label, f) in enumerate([
    ("Square root", "=SQRT(I35)"),
    ("Squared", "=I35^2"),
    ("Even?", "=IF(ISEVEN(I35),\"yes\",\"no\")"),
    ("As a percentage of 1,000", '=TEXT(I35/1000,"0.0%")'),
    ("Factorial of 10", "=FACT(10)"),
    ("Pi to 6 places", "=ROUND(PI(),6)"),
    ("Random 1-100 (F9 to reroll)", "=RANDBETWEEN(1,100)"),
    ("Days until New Year", "=DATE(YEAR(TODAY())+1,1,1)-TODAY()"),
    ("Week of the year", "=WEEKNUM(TODAY())"),
    ("18% tip on $86.40", "=ROUND(86.4*0.18,2)"),
    ("...split 4 ways", "=ROUND(86.4*1.18/4,2)"),
    ("22 C in Fahrenheit", "=22*9/5+32"),
    ("Marathon in km", "=ROUND(26.2*1.609344,2)"),
], start=36):
    put(f"H{r}", label)
    put(f"I{r}", f)

# ---------------------------------------------------------------- summary
put("A52", "THE BIG PICTURE")
put("A53", "Housing cost as % of pay")
put("B53", '=TEXT(B13/B35,"0.0%")')
put("C53", '=IF(B13/B35>0.28,"above the 28% guideline","within the 28% guideline")')
put("A54", "Savings in 25 years (today's money)")
put("B54", "=B29")
put("A55", "Years of spending it would cover")
put("B55", "=ROUND(B29/(C45*12),1)")

# ---------------------------------------------------------------- write
max_r = max(int("".join(ch for ch in a if ch.isdigit())) for a in cells)
max_c = max("ABCDEFGHIJ".index(a[0]) for a in cells)
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "finance-calculator.csv")
with open(out, "w", newline="", encoding="utf-8") as fh:
    w = csv.writer(fh)
    for r in range(1, max_r + 1):
        w.writerow([cells.get(f"{col(c)}{r}", "") for c in range(max_c + 1)])
print(out, len(cells), "cells,", sum(1 for v in cells.values() if v.startswith("=")), "formulas")
