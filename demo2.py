from calendary import Calendary

cal = Calendary(2018)
date_list = []
cal = cal.month(8, work=True, workweek_start=1, workweek_end=5)
for weekday, date in cal:
    date_list.append(str(date))
print(date_list)
