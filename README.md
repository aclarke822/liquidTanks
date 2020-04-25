# liquidTanks
Simple SCADA Tank Report Script

I broke up your script into several functions that each do a job. GetXMLValues, PrepareDictionary, WriteToFile.
Needs testing and error checking but should be easier to maintain than what you had.
Did not modify any other functions except CheckValue to get easier values to deal with.
Changed up the LogLevel implementation to make more sense. Message me if you have any questions.
