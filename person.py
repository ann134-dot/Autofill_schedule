class Person:
    def __init__(self, name, role, schedule) -> None:
        self.name = name
        self.role = role
        self.schedule = schedule

    def count_day(self):
        hours = 0
        days = 0
        for key, val in self.schedule.items():
            if int(val) == 7:
                hours+=11
                days+=1
            elif int(val) == 4:
                if key.day >=30:
                    hours+=1
                else:
                    hours+=4
        return (days, hours)
        
    def count_night(self):
        hours = 0
        days = 0
        for key, val in self.schedule.items():
            if int(val) == 4:
                if key.day >= 30:
                    hours+=2
                else:
                    hours+=7
                days+=1
        return (days, hours)
    
    def count_vacation(self):
        days = 0
        for key, val in self.schedule.items():
            if int(val) == 1:
                days+=1
        return days
    
    def count_total_work(self):
        hours = 0
        days = 0
        day = self.count_day()
        night = self.count_night()
        return (day[0]+night[0], day[1]+night[1])
    