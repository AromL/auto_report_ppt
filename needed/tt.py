import datetime as dt

print((dt.date.today() - dt.timedelta(days=dt.date.today().isocalendar()[2])).strftime('%Y-%m-%d'))

print(dt.date.today().isocalendar())
