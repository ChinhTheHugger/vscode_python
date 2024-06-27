from datetime import datetime
import datetime
import time
msg = "test hello"
print(msg)
dt = datetime.datetime.today()
print(dt)
print(dt.strftime("%d_%m_%Y_%H_%M_%S"))
print(int(datetime.datetime.timestamp(dt)*1000))
dt = datetime.date.today()
print(dt)
print(dt.strftime("%d_%m_%Y_%H_%M_%S"))
print(time.mktime(dt.timetuple()))
print(datetime.date.fromtimestamp(time.mktime(dt.timetuple())))