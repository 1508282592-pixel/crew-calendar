from ics import Calendar, Event
from datetime import datetime, timedelta

c = Calendar()

e = Event()
e.name = "✈️ Crew Calendar Test"
e.begin = datetime.now()
e.duration = timedelta(hours=1)
e.description = "系统测试成功"

c.events.add(e)

with open("crew_schedule.ics", "w") as f:
    f.writelines(c)
