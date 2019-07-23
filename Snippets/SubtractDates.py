from datetime import *

def checkDays(thisRun, lastRun, delay):
    dayCount = (thisRun - lastRun).days #convert to number of days
    if dayCount < delay:
        return False
    return True


def get120DaysAgo():
    longAgo = datetime.today() - timedelta(days = 120)
    return longAgo
