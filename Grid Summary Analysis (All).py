import arcpy, os, datetime
from IPython.display import display
import pandas as pd
import numpy as np
import xlsxwriter

arcpy.management.Delete(r"memory") # delete anything in memory before running script
arcpy.env.workspace = r"memory" # use memory as a workspace for faster performance
arcpy.env.addOutputsToMap = False # don't add outputs to map for faster processing
arcpy.env.overwriteOutput = True # overwrite existing outputs

basePath = r"O:\GIS\Projects\41160-002\ArcGISPro\Clearwater CMOM MML V2\Clearwater CMOM MML V2.gdb" # path holding the base data
gridSummaryPath = os.path.join(basePath, "Grid_Summary") # path to the grid summary

#priorityGridQuery = f"ATLAS_NUMB IN ('286B', '270B', '295B', '260A', '289A', '288B', '258A', '267A', '269B')" # use to query out high priority atlas grids
ssoQuery = f"root_cause = 'Debris/Blockage' Or root_cause = 'Grease' Or root_cause = 'Lateral' And (asset_type = 'Lateral Sewer' Or asset_type = 'Mainline Sewer')" # query to include only the appropriate ssos

from arcgis.gis import GIS

gis = GIS("pro") # log into Portal using ArcPro

sewerMainID = "XXXX" # item ids of feature services we need to use in analysis
CMOM_ID = "XXXX" # contains multiple sublayers
greaseID = "XXXX" # published under my account
gridID = "XXXX"

sewerMainItem = gis.content.get(sewerMainID)
CMOM_Item = gis.content.get(CMOM_ID) # use this single item to call multiple sublayers
greaseItem = gis.content.get(greaseID)
gridItem = gis.content.get(gridID)

mainEndpoint = f"{sewerMainItem.url}/0" 
ssoEndpoint = f"{CMOM_Item.url}/10"
hotspotEndpoint = f"{CMOM_Item.url}/0"
sagEndpoint = f"{CMOM_Item.url}/1"
rootEndpoint = f"{CMOM_Item.url}/2"
greaseEndpoint = f"{greaseItem.url}/0"
gridEndpoint = f"{gridItem.url}/0"

endpointDict = {"mains": mainEndpoint, "sso": ssoEndpoint, "hotspot": hotspotEndpoint, "sag": sagEndpoint, "root": rootEndpoint, "grease": greaseEndpoint, "grid": gridEndpoint} # create dictionary 
display(endpointDict["hotspot"])

copiedFCDict = {} # store the copied feature classes in this dictionary
for x in endpointDict:
    newFCName = x  # grab key name
    endpoint = endpointDict[x] # grab endpoint
    
    fcResult = arcpy.management.CopyFeatures(endpoint, newFCName) # create new feature class from raw data
    
    copiedFCDict[newFCName] = fcResult # add in copied feaure class

finalFLDict = {} # store feature layers in this dictionary
for x in copiedFCDict:
    finalFLName = x
    print("Final Feature Layer name: " + finalFLName)
    copiedFC = copiedFCDict[x] # grab FC in dict
    
    if finalFLName == "grid":
        #finalFL = arcpy.management.MakeFeatureLayer(copiedFC, finalFLName, priorityGridQuery) # create feature layer if grid pass in query for priority grids
        finalFL = arcpy.management.MakeFeatureLayer(copiedFC, finalFLName) # create feature layer if grid pass in query for ALL grids
        
    elif finalFLName == "sso":
        finalFL = arcpy.management.MakeFeatureLayer(copiedFC, finalFLName, ssoQuery) # create feature layer if it's ssos with query
        
    elif finalFLName == "mains":
        query = f"ORIG_LAYER = 'SANITARY' AND DIAMETER >= 4"
        finalFL = arcpy.management.MakeFeatureLayer(copiedFC, finalFLName, query) # create feature layer for mains that are not privately owned
    
    else:
        finalFL = arcpy.management.MakeFeatureLayer(copiedFC, finalFLName) # create feature layer
    
    finalFLDict[finalFLName] = finalFL # add in final feature layer

arcpy.management.AddField(finalFLDict["mains"], "PIPE_GIS_LENGTH", "TEXT", field_length = 30) # add intermediate field for carrying over pipe length
arcpy.management.CalculateField(finalFLDict["mains"], "PIPE_GIS_LENGTH", "!shape.length@FEET!") # carry over shape length data
    
arcpy.management.DeleteField(finalFLDict["grid"], "ATLAS_NUMB", "KEEP_FIELDS") # keep only the name field in the grid feature layer, delete others
arcpy.management.AddField(finalFLDict["grid"], "SSO_COUNT", "SHORT") # add fields for variables
arcpy.management.AddField(finalFLDict["grid"], "SSO_SCORE", "SHORT")
arcpy.management.AddField(finalFLDict["grid"], "HOTSPOT_COUNT", "SHORT")
arcpy.management.AddField(finalFLDict["grid"], "HOTSPOT_SCORE", "SHORT")
arcpy.management.AddField(finalFLDict["grid"], "SAG_COUNT", "SHORT")
arcpy.management.AddField(finalFLDict["grid"], "SAG_SCORE", "SHORT")
arcpy.management.AddField(finalFLDict["grid"], "GREASE_COUNT", "SHORT")
arcpy.management.AddField(finalFLDict["grid"], "GREASE_SCORE", "SHORT")
arcpy.management.AddField(finalFLDict["grid"], "ROOTS_COUNT", "SHORT")
arcpy.management.AddField(finalFLDict["grid"], "ROOTS_SCORE", "SHORT")
arcpy.management.AddField(finalFLDict["grid"], "TOTAL_SCORE", "SHORT")

ssoMulti = 5 # scoring for each variable
hotspotMulti = 4
sagMulti = 3
greaseMulti = 3
rootsMulti = 2

ssoJoin = arcpy.analysis.SpatialJoin(finalFLDict["grid"], finalFLDict["sso"], "ssoJoin", match_option = "INTERSECT") # spatially join the SSOs to the grids
arcpy.management.AddJoin(finalFLDict["grid"], "OBJECTID", ssoJoin, "TARGET_FID") # join the data back up

expression = "!ssoJoin.Join_Count!"
arcpy.management.CalculateField(finalFLDict["grid"], "grid.SSO_COUNT", expression) # field calc the number of ssos in grid cell
arcpy.management.RemoveJoin(finalFLDict["grid"]) # remove the join

expression = f"!SSO_COUNT! * {ssoMulti}"
arcpy.management.CalculateField(finalFLDict["grid"], "SSO_SCORE", expression) # calc the total score for ssos for grid

arcpy.management.Delete(ssoJoin)

hotspotTemp = arcpy.management.FeatureToPoint(finalFLDict["hotspot"], "hotspotTemp", "INSIDE") # convert line to point, use the midpoint of each main
hotspotJoin = arcpy.analysis.SpatialJoin(finalFLDict["grid"], hotspotTemp, "hotspotJoin", match_option = "INTERSECT") # spatially join the hotspots to the grids
arcpy.management.AddJoin(finalFLDict["grid"], "OBJECTID", hotspotJoin, "TARGET_FID") # join the data back up

expression = "!hotspotJoin.Join_Count!"
arcpy.management.CalculateField(finalFLDict["grid"], "grid.HOTSPOT_COUNT", expression) # field calc the number of hotspots in grid cell
arcpy.management.RemoveJoin(finalFLDict["grid"]) # remove the join

expression = f"!HOTSPOT_COUNT! * {hotspotMulti}"
arcpy.management.CalculateField(finalFLDict["grid"], "HOTSPOT_SCORE", expression) # calc the total score for hotspots for grid

arcpy.management.Delete(hotspotJoin)

sagTemp = arcpy.management.FeatureToPoint(finalFLDict["sag"], "sagTemp", "INSIDE") # convert line to point, use the midpoint of each main
sagJoin = arcpy.analysis.SpatialJoin(finalFLDict["grid"], sagTemp, "sagJoin", match_option = "INTERSECT") # spatially join the sags to the grids
arcpy.management.AddJoin(finalFLDict["grid"], "OBJECTID", sagJoin, "TARGET_FID") # join the data back up

expression = "!sagJoin.Join_Count!"
arcpy.management.CalculateField(finalFLDict["grid"], "grid.SAG_COUNT", expression) # field calc the number of sags in grid cell
arcpy.management.RemoveJoin(finalFLDict["grid"]) # remove the join

expression = f"!SAG_COUNT! * {sagMulti}"
arcpy.management.CalculateField(finalFLDict["grid"], "SAG_SCORE", expression) # calc the total score for sags for grid

arcpy.management.Delete(sagJoin)

greaseTemp = arcpy.management.FeatureToPoint(finalFLDict["grease"], "greaseTemp", "INSIDE") # convert line to point, use the midpoint of each main
greaseJoin = arcpy.analysis.SpatialJoin(finalFLDict["grid"], greaseTemp, "greaseJoin", match_option = "INTERSECT") # spatially join the grease to the grids
arcpy.management.AddJoin(finalFLDict["grid"], "OBJECTID", greaseJoin, "TARGET_FID") # join the data back up

expression = "!greaseJoin.Join_Count!"
arcpy.management.CalculateField(finalFLDict["grid"], "grid.GREASE_COUNT", expression) # field calc the number of grease in grid cell
arcpy.management.RemoveJoin(finalFLDict["grid"]) # remove the join

expression = f"!GREASE_COUNT! * {greaseMulti}"
arcpy.management.CalculateField(finalFLDict["grid"], "GREASE_SCORE", expression) # calc the total score for grease for grid

arcpy.management.Delete(greaseJoin)

rootTemp = arcpy.management.FeatureToPoint(finalFLDict["root"], "rootTemp", "INSIDE") # convert line to point, use the midpoint of each main
rootJoin = arcpy.analysis.SpatialJoin(finalFLDict["grid"], rootTemp, "rootJoin", match_option = "INTERSECT") # spatially join the roots to the grids
arcpy.management.AddJoin(finalFLDict["grid"], "OBJECTID", rootJoin, "TARGET_FID") # join the data back up

expression = "!rootJoin.Join_Count!"
arcpy.management.CalculateField(finalFLDict["grid"], "grid.ROOTS_COUNT", expression) # field calc the number of roots in grid cell
arcpy.management.RemoveJoin(finalFLDict["grid"]) # remove the join

expression = f"!ROOTS_COUNT! * {rootsMulti}"
arcpy.management.CalculateField(finalFLDict["grid"], "ROOTS_SCORE", expression) # calc the total score for roots for grid

arcpy.management.Delete(rootJoin)

expression = "!SSO_SCORE! + !HOTSPOT_SCORE! + !SAG_SCORE! + !GREASE_SCORE! + !ROOTS_SCORE!" # sum up all scores
arcpy.management.CalculateField(finalFLDict["grid"], "TOTAL_SCORE", expression)
gridSummary = arcpy.conversion.ExportFeatures(finalFLDict["grid"], gridSummaryPath) # export summary results

print("Constructing feature layer for each grid.")

gridList = [row[0] for row in arcpy.da.SearchCursor(finalFLDict["grid"], "ATLAS_NUMB")] # generate entire list of grids
gridList.sort() #sort by name ascending
gridFLDictionary = {} # create a dictionary for each feature layer for each grid cell
endpoint = endpointDict["grid"] # grab the endpoint for grid
grid_copy = arcpy.management.CopyFeatures(endpoint, "grid_copy") # make a local copy of the grid from endpoint

for grid in gridList: # create a separate feature layer for each grid cell
    expression = "ATLAS_NUMB = '{}'".format(grid) # query out each grid cell 
    gridFL = arcpy.management.MakeFeatureLayer(grid_copy, grid + "_FL", expression) # create feature layer for single grid cell
    print("Current grid: " + grid)
    gridFLDictionary[grid] = gridFL # store key value pair to dictionary
    
    arcpy.management.DeleteField(gridFLDictionary[grid], "ATLAS_NUMB", "KEEP_FIELDS") # keep only the name field in the grid feature layer, delete others
    arcpy.management.AddField(gridFLDictionary[grid], "PIPENAME", "TEXT", field_length = 30) # add fields for pipe CMOM concerns
    arcpy.management.AddField(gridFLDictionary[grid], "MHUP", "TEXT", field_length = 30)
    arcpy.management.AddField(gridFLDictionary[grid], "MHDN", "TEXT", field_length = 30)
    arcpy.management.AddField(gridFLDictionary[grid], "DIAMETER", "DOUBLE")
    arcpy.management.AddField(gridFLDictionary[grid], "Hazen_Material_V2", "TEXT", field_length = 30)
    arcpy.management.AddField(gridFLDictionary[grid], "CCTV", "TEXT", field_length = 30)
    arcpy.management.AddField(gridFLDictionary[grid], "CCTV_DATE", "DATE")
    arcpy.management.AddField(gridFLDictionary[grid], "LINER_TYPE", "TEXT", field_length = 30)
    arcpy.management.AddField(gridFLDictionary[grid], "LINER_DATE", "DATE")
    arcpy.management.AddField(gridFLDictionary[grid], "HAS_HOTSPOT", "TEXT", field_length = 30)
    arcpy.management.AddField(gridFLDictionary[grid], "SSO_DISTANCE_HOTSPOT", "DOUBLE")
    arcpy.management.AddField(gridFLDictionary[grid], "GREASE_VALUE", "TEXT", field_length = 30)
    arcpy.management.AddField(gridFLDictionary[grid], "GREASE_AMOUNT", "TEXT", field_length = 30)
    arcpy.management.AddField(gridFLDictionary[grid], "HAS_SAG", "TEXT", field_length = 30)
    arcpy.management.AddField(gridFLDictionary[grid], "SAG_PER", "TEXT", field_length = 30)
    arcpy.management.AddField(gridFLDictionary[grid], "HAS_ROOTS", "TEXT", field_length = 30)
    arcpy.management.AddField(gridFLDictionary[grid], "PIPE_GIS_LENGTH", "TEXT", field_length = 30)
    

arcpy.management.SelectLayerByAttribute(finalFLDict["sso"], "CLEAR_SELECTION") # clear selections
arcpy.management.SelectLayerByAttribute(finalFLDict["hotspot"], "CLEAR_SELECTION") # clear selections

arcpy.management.AddField(finalFLDict["hotspot"], "SSO_DISTANCE_HOTSPOT", "DOUBLE") # add field to indicate distance to SSO
nearTable = arcpy.analysis.GenerateNearTable(finalFLDict["hotspot"], finalFLDict["sso"], "nearTable", "300", distance_unit = "Feet") # find the distance to the closest sso
arcpy.conversion.ExportTable(nearTable, os.path.join(basePath, "nearTableHotspotSSO"))                       

arcpy.management.AddJoin(finalFLDict["hotspot"], "OBJECTID", nearTable, "IN_FID") # join the data

expression = "nearTable.NEAR_DIST <= 300"
selection = arcpy.management.SelectLayerByAttribute(finalFLDict["hotspot"], "NEW_SELECTION", expression) # select rows that have distances <= 200 to a SSO
print("Total SSOs: " + str(arcpy.management.GetCount(selection)))
arcpy.management.CalculateField(finalFLDict["hotspot"], "hotspot.SSO_DISTANCE_HOTSPOT", "!nearTable.NEAR_DIST!") # field calc over the distance to the closest SSO

arcpy.management.RemoveJoin(finalFLDict["hotspot"])    
arcpy.management.SelectLayerByAttribute(finalFLDict["hotspot"], "CLEAR_SELECTION")

print("Converting mains into points.")

pointFLDict = {} # convert each main to a point to aid in spatial join
for x in finalFLDict:
    finalFLName = x + "_Point"
    
    if finalFLName == "grid" or finalFLName == "sso": # skip over grid and ssos
        continue
        
    else:
        point = arcpy.management.FeatureToPoint(finalFLDict[x], finalFLName, "INSIDE") # convert line to point, use the midpoint of each main 
    
    pointFLDict[finalFLName] = point # add in converted point

print("Appending in main data into each grid.")

gridTableDict = {} # convert the spatial format for each grid cell into tabular form
mains_pointFL = arcpy.management.MakeFeatureLayer(pointFLDict["mains_Point"], "mains_pointFL") # create feature layer for point mains
  
for x in gridFLDictionary:
    print("Current Dictionary Key: " + x)
    arcpy.management.SelectLayerByLocation(mains_pointFL, "INTERSECT", gridFLDictionary[x]) # select point mains that intersect each grid
    Temp_Points = arcpy.management.CopyFeatures(mains_pointFL, "Temp_Points") # create temp layer for intersecting points
    gridTable = arcpy.conversion.ExportTable(gridFLDictionary[x], "gridTable") # convert grid to table, need to append in main info so we can join up data later
    mainTable = arcpy.conversion.ExportTable(Temp_Points, "mainTable") # convert intersecting points to table, add to dictionary under grid name
    
    arcpy.management.Append(mainTable, gridTable, "NO_TEST") # append in main data to the grid table so we can join data to it

    name = "Grid_" + x
    outputPath = os.path.join(basePath, name)
    gridTableExport = arcpy.conversion.ExportTable(gridTable, outputPath) # export grid table with appended mains to gdb, will edit these tables later
    gridTableDict[x] = gridTableExport   # save exported grid table to dictionary to later create a table view
    
    arcpy.management.SelectLayerByAttribute(mains_pointFL, "CLEAR_SELECTION") # clear the selection                                
    arcpy.management.Delete(gridTable) # memory management
    arcpy.management.Delete(mainTable)
    arcpy.management.Delete(Temp_Points)

print("Creating Table Views for each grid.")

gridTableTVDict = {} # store the table views in a dictionary

for x in gridTableDict:
    gridTableTVDict[x] = arcpy.management.MakeTableView(gridTableDict[x], x) # create table view

print("Updating liner data.")

for x in gridTableTVDict:
  
    arcpy.management.AddJoin(gridTableTVDict[x], "PIPENAME", finalFLDict["mains"], "PIPENAME") # join up main data to each grid table
    expression = f"(mains.HZ_RehabAction = 'Already Lined' or mains.HZ_RehabAction = 'CIPP') And mains.LINER_TYPE IS NULL" # select the fields regarding work being done
    arcpy.management.SelectLayerByAttribute(gridTableTVDict[x], "NEW_SELECTION", expression) # select valid joins
    arcpy.management.CalculateField(gridTableTVDict[x], "LINER_TYPE", "\'Lined\'") # update roots
    
    arcpy.management.RemoveJoin(gridTableTVDict[x])
    
    arcpy.management.SelectLayerByAttribute(gridTableTVDict[x], "CLEAR_SELECTION")

print("Adding hotspot data.")

for x in gridTableTVDict:
    arcpy.management.AddJoin(gridTableTVDict[x], "PIPENAME", finalFLDict["hotspot"], "pipename") # join up hotspot data to each grid table

    expression = "hotspot.pipename IS NOT NULL"  
    arcpy.management.SelectLayerByAttribute(gridTableTVDict[x], "NEW_SELECTION", expression) # select valid joins
    arcpy.management.CalculateField(gridTableTVDict[x], "HAS_HOTSPOT", "'{}'".format("Yes")) # update to be yes to indicate hotspot
    
    expression = "hotspot.SSO_DISTANCE_HOTSPOT IS NOT NULL" # select where distance is not null
    selection = arcpy.management.SelectLayerByAttribute(gridTableTVDict[x], "NEW_SELECTION", expression) # select valid sso distances
    
    if arcpy.management.GetCount(selection) != 0: # inform where ssos happened
        print("Total SSOs: " + str(arcpy.management.GetCount(selection)) + f" on Grid {x}")
    
    arcpy.management.CalculateField(gridTableTVDict[x], "SSO_DISTANCE_HOTSPOT", "!hotspot.SSO_DISTANCE_HOTSPOT!") # update to carry over distance
    
    arcpy.management.RemoveJoin(gridTableTVDict[x])
    
    arcpy.management.SelectLayerByAttribute(gridTableTVDict[x], "SWITCH_SELECTION") # switch selection to update remaining to be no
    arcpy.management.CalculateField(gridTableTVDict[x], "HAS_HOTSPOT", f"'{No}'") # update to be no
    
    arcpy.management.SelectLayerByAttribute(gridTableTVDict[x], "CLEAR_SELECTION")
    

print("Adding grease data.")

for x in gridTableTVDict:
    arcpy.management.AddJoin(gridTableTVDict[x], "PIPENAME", finalFLDict["grease"], "pipename") # join up grease data to each grid table
    
    expression = "grease.pipename IS NOT NULL"
    arcpy.management.SelectLayerByAttribute(gridTableTVDict[x], "NEW_SELECTION", expression) # select valid joins
    arcpy.management.CalculateField(gridTableTVDict[x], "GREASE_VALUE", "!grease.GREASE_AMOUNT!") # update grease amount
    arcpy.management.CalculateField(gridTableTVDict[x], "GREASE_AMOUNT", "!grease.GREASE_VALUE!") # update grease value
    
    arcpy.management.RemoveJoin(gridTableTVDict[x])
    
    arcpy.management.SelectLayerByAttribute(gridTableTVDict[x], "SWITCH_SELECTION") # switch selection to update remaining grease entries to be none
    arcpy.management.CalculateField(gridTableTVDict[x], "GREASE_VALUE", f"'None'") # indicate no grease
    arcpy.management.CalculateField(gridTableTVDict[x], "GREASE_AMOUNT", f"'None'") # indicate no grease
    
    arcpy.management.SelectLayerByAttribute(gridTableTVDict[x], "CLEAR_SELECTION")


print("Adding sag data.")

for x in gridTableTVDict:
    arcpy.management.AddJoin(gridTableTVDict[x], "PIPENAME", finalFLDict["sag"], "pipename") # join up sag data to each grid table
    expression = "sag.pipename IS NOT NULL"
    arcpy.management.SelectLayerByAttribute(gridTableTVDict[x], "NEW_SELECTION", expression) # select valid joins
    arcpy.management.CalculateField(gridTableTVDict[x], "HAS_SAG", f"'Yes'") # update sag amount
    arcpy.management.CalculateField(gridTableTVDict[x], "SAG_PER", "!sag.sag_percent!") # update sag value
    
    arcpy.management.RemoveJoin(gridTableTVDict[x])
    
    arcpy.management.SelectLayerByAttribute(gridTableTVDict[x], "SWITCH_SELECTION") # switch selection to update remaining sag entries to be none
    arcpy.management.CalculateField(gridTableTVDict[x], "HAS_SAG", f"'No'") # indicate no sag
    arcpy.management.CalculateField(gridTableTVDict[x], "SAG_PER", f"'None'") # indicate no sag
    
    arcpy.management.SelectLayerByAttribute(gridTableTVDict[x], "CLEAR_SELECTION")
    

print("Adding root data.")

for x in gridTableTVDict:
    arcpy.management.AddJoin(gridTableTVDict[x], "PIPENAME", finalFLDict["root"], "pipename") # join up root data to each grid table
    expression = "root.pipename IS NOT NULL"
    arcpy.management.SelectLayerByAttribute(gridTableTVDict[x], "NEW_SELECTION", expression) # select valid joins
    arcpy.management.CalculateField(gridTableTVDict[x], "HAS_ROOTS", f"'Yes'") # update roots
    
    arcpy.management.RemoveJoin(gridTableTVDict[x])
    
    arcpy.management.SelectLayerByAttribute(gridTableTVDict[x], "SWITCH_SELECTION") # switch selection to update remaining root entries to be none
    arcpy.management.CalculateField(gridTableTVDict[x], "HAS_ROOTS", f"'No'") # indicate no roots
    
    arcpy.management.SelectLayerByAttribute(gridTableTVDict[x], "CLEAR_SELECTION")
    

print("Converting dBASE tables to Pandas Dataframes.")

def arcgis_table_to_dataframe(in_fc, input_fields, query="", skip_nulls=False, null_values=None):
    """Function will convert an arcgis table into a pandas dataframe with an object ID index, and the selected
    input fields. Uses TableToNumPyArray to get initial data.
    :param - in_fc - input feature class or table to convert
    :param - input_fields - fields to input into a da numpy converter function
    :param - query - sql like query to filter out records returned
    :param - skip_nulls - skip rows with null values
    :param - null_values - values to replace null values with.
    :returns - pandas dataframe
    source: https://gis.stackexchange.com/questions/450852/read-a-gdb-table-into-a-pandas-dataframe-in-arcpy#:~:text=def%20arcgis_table_to_dataframe(,input_fields)%0A%20%20%20%20return%20fc_dataframe
    """
    OIDFieldName = arcpy.Describe(in_fc).OIDFieldName
    if input_fields:
        final_fields = [OIDFieldName] + input_fields
    else:
        final_fields = [field.name for field in arcpy.ListFields(in_fc)]
    np_array = arcpy.da.TableToNumPyArray(in_fc, final_fields, query, skip_nulls, null_values)
    object_id_index = np_array[OIDFieldName]
    fc_dataframe = pd.DataFrame(np_array, index=object_id_index, columns = input_fields, dtype = str)
    return fc_dataframe

columns = [f.name for f in arcpy.ListFields(gridSummary) if f.type!="Geometry"] #List the fields you want to include. I want all columns except the geometry
gridSummaryDF = pd.DataFrame(data=arcpy.da.SearchCursor(gridSummary, columns), columns = columns, dtype = str) # pass in grid summary
gridSummaryDF.drop(axis = 1, columns = ['OBJECTID', 'Shape_Length', 'Shape_Area'], inplace = True) # drop these fields, not needed

gridSummaryDF["SUM_LENGTH_NO_CCTV"] = np.nan
gridSummaryDF["TOTAL_LENGTH_SEWER_LENGTH"] = np.nan 
gridSummaryDF["PERCENTAGE_NO_CCTV"] = np.nan
gridSummaryDF.sort_values(by=["ATLAS_NUMB"], inplace = True) # sort table by Atlas Number, replace existing dataframe with sorted DF

dataFrameDict = {} # store the pandas frames in a dict

for x in gridTableDict:
    df = arcgis_table_to_dataframe(gridTableDict[x], None) # object id is used as the index
    removedRowDF = df.drop(1) # drop the first index, it was made way back in the creation of the grid indices
    gridIDDF = removedRowDF.assign(ATLAS_NUMB = x) # set Atlas Grid ID value to column
    objectIDDF = gridIDDF.drop(axis = 1, columns = 'OBJECTID') # drop object ID, not needed
    dataFrameDict[x] = objectIDDF # add to dictionary
    
print("Updating Pandas Dataframes.")
for x in dataFrameDict: # iterate through each data frame for each grid
    dataFrameDict[x]['PIPE_GIS_LENGTH'] = dataFrameDict[x]['PIPE_GIS_LENGTH'].astype(float) # convert the GIS_PIPE_LENGTH field to float
    dataFrameDict[x]['SSO_DISTANCE_HOTSPOT'] = dataFrameDict[x]['SSO_DISTANCE_HOTSPOT'].astype(float) # convert the SSO_DISTANCE_HOTSPOT field to float
    
    noCCTVSum = dataFrameDict[x].loc[dataFrameDict[x]['CCTV'] != 'Yes', 'PIPE_GIS_LENGTH'].sum() # sum up all the pipes that have been not been CCTV'd
    totalLength = dataFrameDict[x]['PIPE_GIS_LENGTH'].sum() # sum up all the pipes
    
    try:
        percentage = float(noCCTVSum) / float(totalLength) * 100 # calculate percenatage
    except:
        percentage = 0
        print("Possibly no mains within grid.")
        
    gridSummaryDF.loc[gridSummaryDF['ATLAS_NUMB'] == x, 'SUM_LENGTH_NO_CCTV'] = noCCTVSum # update the grid summary value for no cctv
    gridSummaryDF.loc[gridSummaryDF['ATLAS_NUMB'] == x, 'TOTAL_LENGTH_SEWER_LENGTH'] = totalLength # update the grid summary value for total sewer length
    gridSummaryDF.loc[gridSummaryDF['ATLAS_NUMB'] == x, 'PERCENTAGE_NO_CCTV'] = percentage # update the grid summary value for percentage cctv'd
    
    gridSummaryDF = gridSummaryDF.round(2) # round each column to 2 decimals
    dataFrameDict[x] = dataFrameDict[x].round(2) # round each column to 2 decimals

print("Exporting dataframes to xlsx.")

folder = r"O:\GIS\Projects\41160-002\Data\Tables\Grid Prioritization"
xlsx = "Grid_All_Prioritization.xlsx"
outputxlsx = os.path.join(folder, xlsx) # path to the final xlsx

with pd.ExcelWriter(outputxlsx, engine = "xlsxwriter") as writer:
    workbook = writer.book # create workbook object from Pandas Writer
    align_format = workbook.add_format({'align': 'center', 'valign': 'center'}) # create format object in Workbook, use it to write the style to each cell
    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd', 'align': 'center', 'valign': 'center'})
    
    gridSummaryDF.to_excel(writer, sheet_name = 'Grid Summary') # write to the xlsx file
    
    for column in gridSummaryDF:
        column_width = max(gridSummaryDF[column].astype(str).map(len).max(), len(column)) + 5 # grab longest row value in table, add 5 to it, will be column width
        col_idx = gridSummaryDF.columns.get_loc(column) + 1 # adding one corrects the indexing
        writer.sheets["Grid Summary"].set_column(col_idx, col_idx, column_width) # update column widths
    
    for row_num, row_data in enumerate(gridSummaryDF.values, start=1): # do this to center text
        for col_num, cell_value in enumerate(row_data):
            writer.sheets["Grid Summary"].write(row_num, col_num + 1, cell_value, align_format) # add one, consider Index
    
    for x in dataFrameDict:
        dataFrameDict[x].to_excel(writer, sheet_name = x) # loop through the dataframes, write to excel file
        
        for column in dataFrameDict[x]:
            column_width = max(dataFrameDict[x][column].astype(str).map(len).max(), len(column)) + 5 # grab longest row value in table, add 5 to it, will be column width
            col_idx = dataFrameDict[x].columns.get_loc(column) + 1
            writer.sheets[x].set_column(col_idx, col_idx, column_width) # update column widths
            
        for row_num, row_data in enumerate(dataFrameDict[x].values, start=1): # do this to center text
            for col_num, cell_value in enumerate(row_data):
                writer.sheets[x].write(row_num, col_num + 1, str(cell_value), align_format) # add one, consider Index column
