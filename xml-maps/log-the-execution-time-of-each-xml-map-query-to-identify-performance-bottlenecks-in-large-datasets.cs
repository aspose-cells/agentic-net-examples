// Title: Measure and log execution time of each XmlMap.ImportXMLData operation in a large Excel workbook using Aspose.Cells for .NET
// AI Prompts: Wrap each XmlMap.ImportXMLData call in a Stopwatch, capture the elapsed milliseconds, and write a formatted line to a log file for that map. | After processing all XML maps, calculate the total import duration and append a summary entry to the same log file. | Implement a reusable method that receives a Workbook and returns a Dictionary<string, long> mapping XML map names to their individual import times in milliseconds.
// Common Searches: how to log import time for each XML map in Aspose.Cells C# | benchmark XmlMap.ImportXMLData performance in large Excel files | measure execution duration of XML map queries with Aspose.Cells .NET | profile XML map import speed for multiple maps using Stopwatch
// Tags: Aspose.Cells XmlMap performance measurement | C# Stopwatch XML map duration | log XML map import latency | profiling Aspose.Cells XML map queries | measure XmlMap.ImportXMLData execution

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through its XmlMaps collection, imports data from corresponding XML files while timing each ImportXMLData call with a Stopwatch, logs each map's name and elapsed milliseconds to XmlMapQueryLog.txt, optionally aggregates total time, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output workbook paths
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input workbook not found: {inputPath}");
                return;
            }

            // Load the workbook
            dynamic workbook = new Workbook(inputPath);

            // Path for the execution time log
            const string logFilePath = "XmlMapQueryLog.txt";

            // Open a StreamWriter to record the timings
            using (StreamWriter logWriter = new StreamWriter(logFilePath, false))
            {
                // Attempt to retrieve the XmlMaps collection; if unavailable, skip processing
                if (workbook.XmlMaps == null)
                {
                    Console.WriteLine("The loaded workbook does not contain any XML maps.");
                    return;
                }

                dynamic xmlMaps = workbook.XmlMaps;
                int mapCount = xmlMaps.Count;

                // Iterate through each XML map defined in the workbook
                for (int mapIndex = 0; mapIndex < mapCount; mapIndex++)
                {
                    try
                    {
                        dynamic xmlMap = xmlMaps[mapIndex];

                        // Example XML data source for the map (adjust the path as needed)
                        string xmlDataPath = $"Data{mapIndex}.xml";

                        // Verify that the XML data file exists
                        if (!File.Exists(xmlDataPath))
                        {
                            string missingMsg = $"XML data file not found: {xmlDataPath}";
                            Console.WriteLine(missingMsg);
                            logWriter.WriteLine(missingMsg);
                            continue;
                        }

                        // Measure the time taken to execute the XML map query (ImportXMLData)
                        Stopwatch sw = Stopwatch.StartNew();

                        // Perform the XML map operation
                        xmlMap.ImportXMLData(xmlDataPath);

                        sw.Stop();

                        // Log the execution time
                        string logMessage = $"XML Map '{xmlMap.Name}' (Index {mapIndex}) import time: {sw.ElapsedMilliseconds} ms";
                        Console.WriteLine(logMessage);
                        logWriter.WriteLine(logMessage);
                    }
                    catch (Exception exMap)
                    {
                        string errMsg = $"Error processing XML map at index {mapIndex}: {exMap.Message}";
                        Console.WriteLine(errMsg);
                        logWriter.WriteLine(errMsg);
                    }
                }
            }

            // Save the workbook after processing (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
