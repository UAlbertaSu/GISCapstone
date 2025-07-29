import openpyxl
from openpyxl import load_workbook
import sys # Delete in final
import requests
from datetime import datetime, timedelta, timezone
import time
import json
import os
from dotenv import load_dotenv


def get_next_monday_8am():
    now = datetime.now(timezone.utc)
    days_until_monday = (7 - now.weekday()) % 7
    next_monday = now + timedelta(days=days_until_monday)
    nextMonday_MT = next_monday.astimezone()
    return nextMonday_MT.replace(hour=8, minute=0, second=0, microsecond=0).isoformat()

def get_distance(origin, destination):
    
    load_dotenv()
    api_key = os.getenv("API_KEY")

    # API Config
    
    url = "https://routes.googleapis.com/distanceMatrix/v2:computeRouteMatrix"

    # Request Headers
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": "originIndex,destinationIndex,duration,distanceMeters,status"
    }

    # Request Body with Variables
    payload = {
        "origins": [{
            "waypoint": {
                "address": origin  # Using origin variable
            }
        }],
        "destinations": [{
            "waypoint": {
                "address": destination  # Using destination variable
            }
        }],
        "travelMode": "TRANSIT",
        "transitPreferences": {
            "routingPreference": "LESS_WALKING",
            "allowedTravelModes": ["BUS", "RAIL"]
        },
        "departureTime": get_next_monday_8am()
    }

    # Send Request
    response = requests.post(url, headers=headers, data=json.dumps(payload))
    print(response.json())
    return response.json()

def processSheet(wb, dict):
    # parse the student ID, city and their address in dictionary
    
    sheet = wb.active
    
    
    for row_num, row in enumerate(sheet.iter_rows(values_only = True), start = 1):
        if row_num == 1:
            field = row
        else:
            if field[0] == "Student ID#":
                dict[row[0]] = row[1:3]
                
            else:
                hostList = list(row[:2])
                hostList.append(row[5])
                hostList.append(row[8])
                dict[row[3]] = hostList
    
    return dict
    
def main():
    # Delete this check in MVP
    if sys.platform == "win32":
        studentWb = load_workbook("C:\\Users\steve\SAIT\Sem2\\Capstone\\Inputs\Students\\ModifiedClassList.xlsx")
        hostListWb = load_workbook("C:\\Users\steve\SAIT\Sem2\\Capstone\\Inputs\Locations\\ModifiedHostList.xlsx")
    else:
        studentWb = load_workbook("/Users/stevesu/GIS School/GISCapstone/ModifiedClassList.xlsx")
        hostListWb = load_workbook("/Users/stevesu/GIS School/GISCapstone/ModifiedHost.xlsx")
        
    studentDict = {}
    hostDict = {}
    
    # parse the origins (student address) 
    print("Student Location \n")
    studentDict = processSheet(studentWb, studentDict)
    print(studentDict)
    print('\n')
    print("Host Location \n")
    
    # parse the destination and attributes
    hostDict = processSheet(hostListWb, hostDict)
    print(hostDict)
    print('\n')
    
    # call API
    for k,v in studentDict.items():
        origin = (v[0]+" "+ v[1])
        for key, value in hostDict.items():
            destination = (key+" "+value[2])
            print(origin, destination)
            hostCity = destination[-1]
            studentCity = origin[-1]
            if studentCity.lower() != hostCity.lower():
                continue
            distance = get_distance(origin, destination)
            print(distance)
            exit()

main()