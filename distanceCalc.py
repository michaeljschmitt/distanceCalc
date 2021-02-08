import geopy, openpyxl
from geopy.distance import great_circle

wb = openpyxl.load_workbook("Data.xlsx")
ws = wb.active

wbOutput = openpyxl.Workbook()
wsOutput = wbOutput.active

pointACoordinates = []
pointBCoordinates = []
distances = []

pointATitle = ws["A1"].value
pointBTitle = ws["D1"].value

for row in ws.iter_rows(min_row=2, min_col=1, max_col=3):
	
	pointAID = row[0].value
	latitude = row[1].value
	longitude = row[2].value
	cood = (latitude, longitude)
	pointACoordinates.append([pointAID, cood])

for row in ws.iter_rows(min_row=2, min_col=4, max_col=6):
	
	pointBID = row[0].value
	latitude = row[1].value
	longitude = row[2].value
	cood = (latitude, longitude)
	pointBCoordinates.append([pointBID, cood])

for pointA in pointACoordinates:
	for pointB in pointBCoordinates:
		distance = great_circle(pointA[1],pointB[1]).miles
		distances.append([pointA[0], pointB[0], distance])

shortestpointABDistance = {}
for pointABDistance in distances:
	if pointABDistance[0] == None:
		continue
	elif pointABDistance[0] not in shortestpointABDistance.keys():
		shortestpointABDistance[pointABDistance[0]] = [pointABDistance[1], pointABDistance[2]]
	else:
		if pointABDistance[2] < shortestpointABDistance[pointABDistance[0]][1]:
			shortestpointABDistance[pointABDistance[0]] = [pointABDistance[1], pointABDistance[2]]

#convert dictionary to list
closestDistancesList = []
for key, value in shortestpointABDistance.items():
	closestDistancesList.append([key,value[0],value[1]])
	
wsOutput.append([pointATitle, pointBTitle, 'Distance (mi)'])

# for closest distances (to point A coordinates)
for pointABDistance in closestDistancesList:
    wsOutput.append([pointABDistance[0], pointABDistance[1], pointABDistance[2]])

# for all distances
for pointABDistance in distances:
	wsOutput.append([pointABDistance[0], pointABDistance[1], pointABDistance[2]])

wbOutput.save('./Output.xlsx')