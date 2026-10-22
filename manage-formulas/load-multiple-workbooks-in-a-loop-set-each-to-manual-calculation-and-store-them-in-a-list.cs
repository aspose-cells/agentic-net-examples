// Title: How to load several Excel workbooks, switch each to manual calculation mode, and store them in a List with Aspose.Cells for .NET
// Common Searches: Aspose.Cells load multiple .xlsx files in a loop C# | set workbook calculation mode to manual for each workbook using Aspose.Cells | batch load Excel workbooks and disable automatic calculation with Aspose.Cells | C# reflection to change CalcMode to Manual in Aspose.Cells workbook | store loaded Aspose.Cells Workbook objects in a List collection
// Tags: Aspose.Cells load multiple workbooks | manual calculation mode Aspose.Cells | C# reflection set CalcMode | batch process Excel workbooks Aspose.Cells | store Workbook objects in List

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example shows how to define a list of Excel file paths, iterate through them, verify each file exists, load each workbook with Aspose.Cells, use reflection to set the workbook's calculation mode to Manual when the property is available, add the successfully loaded Workbook objects to a List, handle any loading or reflection errors, and finally report the total number of workbooks loaded.
class Program
{
    static void Main()
    {
        // Define the paths of the workbooks to be loaded.
        var filePaths = new List<string>
        {
            "Book1.xlsx",
            "Book2.xlsx",
            "Book3.xlsx"
            // Add more file names as needed.
        };

        // List that will hold the loaded Workbook objects.
        var workbooks = new List<Workbook>();

        // Loop through each file path, load the workbook, set manual calculation (if supported), and store it.
        foreach (var path in filePaths)
        {
            // Verify that the file exists before attempting to load it.
            if (!File.Exists(path))
            {
                Console.WriteLine($"File not found: {path}");
                continue;
            }

            try
            {
                // Load the workbook from the specified file.
                var wb = new Workbook(path);

                // Attempt to set the calculation mode to Manual using reflection.
                // This avoids compile‑time dependencies on APIs that may not exist in older versions.
                try
                {
                    var settings = wb.Settings;
                    var prop = settings.GetType().GetProperty("CalcMode");
                    if (prop != null && prop.CanWrite)
                    {
                        var enumType = prop.PropertyType;
                        var manualValue = Enum.Parse(enumType, "Manual", ignoreCase: true);
                        prop.SetValue(settings, manualValue);
                    }
                }
                catch (Exception ex)
                {
                    // If reflection fails or the enum value is unavailable, ignore and continue.
                    Console.WriteLine($"Unable to set manual calculation mode for '{path}': {ex.Message}");
                }

                // Add the workbook to the collection.
                workbooks.Add(wb);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook '{path}': {ex.Message}");
            }
        }

        // At this point, 'workbooks' contains all successfully loaded workbooks.
        Console.WriteLine($"Total workbooks loaded: {workbooks.Count}");
    }
}
