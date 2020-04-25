# liquidTanks
Simple SCADA Tank Report Script

I broke up your script into several functions that each do a job. GetXMLValues, PrepareDictionary, WriteToFile.  
Needs testing and error checking but should be easier to maintain than what you had.  
Did not modify any other functions except CheckValue to get easier values to deal with.  
Changed up the LogLevel implementation to make more sense.  
Oh yeah, there is a triple nested dictionary now: Top(Facility(Point(Info))). Message me if you have any questions.  
