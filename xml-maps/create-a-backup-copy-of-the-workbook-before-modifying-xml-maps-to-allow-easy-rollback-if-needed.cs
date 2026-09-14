// Title: Create a backup of an Excel workbook before adding an XML map and setting the worksheet XmlMapIndex with Aspose.Cells for .NET
// AI Prompts: Write C# code that copies an existing .xlsx file to a backup, loads it with Aspose.Cells, adds an XML map from a specified XML file, assigns the map to the first worksheet using the XmlMapIndex property (using reflection if the property is not directly accessible), and saves the modified workbook. | Demonstrate a safe workflow in C# for modifying Excel XML maps with Aspose.Cells: verify source files, create a backup copy, add the XML map, set the worksheet's XmlMapIndex via reflection, and handle any errors gracefully.
// Common Searches: how to backup an Excel file before modifying XML maps using Aspose.Cells C# | c# add xml map to workbook and set worksheet XmlMapIndex with Aspose.Cells | asp.net create copy of .xlsx then associate xml map via reflection | aspose.cells save backup workbook then add xml map example
// Tags: backup workbook before xml map addition Aspose.Cells | add xml map to worksheet Aspose.Cells C# | set worksheet XmlMapIndex via reflection | save modified workbook after xml map Aspose.Cells | validate file existence Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

// The example checks that the original Excel workbook and XML file exist, creates a backup copy of the workbook, loads it with Aspose.Cells, adds an XML map from the XML file, attempts to assign the map to the first worksheet using the XmlMapIndex property via reflection, and finally saves the modified workbook while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            const string originalPath = "original.xlsx";
            const string backupPath = "original_backup.xlsx";
            const string modifiedPath = "original_modified.xlsx";
            const string xmlFilePath = "data.xml";

            // Verify required files exist
            if (!File.Exists(originalPath))
                throw new FileNotFoundException($"Workbook file not found: {originalPath}");
            if (!File.Exists(xmlFilePath))
                throw new FileNotFoundException($"XML file not found: {xmlFilePath}");

            // Load the original workbook
            Workbook workbook = new Workbook(originalPath);

            // Save a backup before modifications
            workbook.Save(backupPath);

            // Add a new XML map from the XML file
            XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;
            int mapIndex = xmlMaps.Add(xmlFilePath);

            // Attempt to associate the first worksheet with the new XML map using reflection
            if (workbook.Worksheets.Count > 0)
            {
                try
                {
                    Worksheet ws = workbook.Worksheets[0];
                    PropertyInfo prop = ws.GetType().GetProperty("XmlMapIndex", BindingFlags.Public | BindingFlags.Instance);
                    if (prop != null && prop.CanWrite)
                    {
                        prop.SetValue(ws, mapIndex);
                    }
                }
                catch (Exception ex)
                {
                    // If association fails, continue without it
                    Console.WriteLine($"Warning: Unable to set XmlMapIndex - {ex.Message}");
                }
            }

            // Save the modified workbook
            workbook.Save(modifiedPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
