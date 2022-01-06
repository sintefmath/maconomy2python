# maconomy2python

follow the instructions at https://stash.code.sintef.no/projects/SPY/repos/sintefpy/browse, which are:
 - git clone ssh://git@git.code.sintef.no/spy/sintefpy.git
 - cd sintefpy/
 - pip install --user .
 - run, e.g.: "~/.local/bin/spy timesheet show", type in your username and password and say yes to saving credentials

**example usage1: **python plotprojectdata.py --projectnumber 123 --totalbudget 1000

example output:

![pie1](pie1.png)
![pie2](pie2.png)
![actuals_per_month](actuals_per_month.png)
![actuals_accumulated](actuals_accumulated.png)

**example usage2: python plotprojectdata.py --filename ExportProjectCard_projectnumber.xlsx --totalbudget 1000**
