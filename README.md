This project contains VBA (Visual Basic for Applications) functions for converting geographical coordinates (latitude and longitude) into UTM (Universal Transverse Mercator) coordinates, identifying latitude bands, and determining grid letters for UTM zones. These functions are useful for mapping and geospatial analysis within Excel or other VBA-compatible environments.

Functions
LLtoUTM
Converts latitude and longitude into UTM eastings and northings.

Parameters:

Lat: Latitude in decimal degrees.
Lon: Longitude in decimal degrees.
ReturnType: A string specifying the return type; use "E" for Easting and "N" for Northing.
Returns: A double representing either the easting or northing of the UTM coordinate, based on the ReturnType parameter.

LatitudeBand
Determines the latitude band letter for a given latitude.

Parameters:

latitude: Latitude in decimal degrees.
Returns: A string representing the latitude band letter.

GridLetter
Calculates the grid letter for a given easting, northing, and UTM zone.

Parameters:

easting: The easting component of the UTM coordinate.
northing: The northing component of the UTM coordinate.
zone: The long integer representing the UTM zone.
Returns: A string composed of two letters representing the UTM grid.
