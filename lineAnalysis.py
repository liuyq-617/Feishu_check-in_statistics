import xlrd
import csv
import datetime
import getopt
import sys
import os
import time
import personalInfo

def StringToMinute(stringIn):
    return int(stringIn[:2]*60 + stringIn[3:])

def evalLeave(personalDict, leaveReason, leaveTime):
    if leaveReason == "事假":
        leaveHour = float(leaveTime[:-2])
        personalDict.alter('businessLeaveHour',  personalDict.get('businessLeaveHour') + leaveHour)
    elif leaveReason == "病假":
        personalDict.alter('sickLeave',  personalDict.get('sickLeave') + 1)
    elif leaveReason == "婚假":
        personalDict.alter('marryLeave',  personalDict.get('marryLeave') + 1)
    elif leaveReason == "丧假":
        personalDict.alter('funeralLeave',  personalDict.get('funeralLeave') + 1)
    elif leaveReason == "陪产假":
        personalDict.alter('maternityLeave',  personalDict.get('maternityLeave') + 1)

def evalFullAttandence(personalDict, checkInTime, workDay):
    if StringToMinute(checkInTime) <= 540 and workDay:
        personalDict.alter('fullAttandenceDay',  personalDict.get('fullAttandenceDay') + 1)

def fullAttandenceAward(personalDict):
    if personalDict.get("fullAttandenceDay") == personalDict.get("attandenceDay"):
        personalDict.alter('fullAttandence', True)

def late(personalDict):
    

def evalSingleLine(personalDict, line):
    workDay = False
    if line[7] == "早9晚6 09:00-18:00;可晚到1小时;可早走1小时":
        personalDict.alter('attandenceDay',  personalDict.get('attandenceDay') + 1)
        workDay = True
    evalLeave(personalDict,line[20],line[19], workDay)
    try:
        evalFullAttandence(personalDict, line[8])
    except BaseException:
        pass

