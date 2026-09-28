// Title: Load several .xlsx files into Aspose.Cells Workbook objects, set each workbook to Manual calculation mode, and collect them in a List in C#
// AI Prompts: Write C# code that iterates over an array of Excel file paths, creates a Workbook for each using Aspose.Cells, sets Settings.CalcMode to Manual, and adds the workbook to a List<Workbook>. | Refactor the loading loop to accept a calculation mode argument (Manual or Automatic) and apply it to each workbook without using reflection. | Enhance the batch loading routine with comprehensive error handling that logs missing files and any exceptions while continuing to load the remaining workbooks.
// Common Searches: aspocells set calculation mode manual for multiple workbooks c# | how to load a list of Excel files into Aspose.Cells and change CalcMode | batch loading Excel workbooks with manual calculation using Aspose.Cells .NET
// Tags: Aspose.Cells batch workbook loading | manual CalcMode configuration | C# workbook collection handling | reflection alternative for CalcMode | robust Excel file loading errors

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using System.Reflection;

// The example demonstrates how to load a predefined collection of .xlsx files into Aspose.Cells Workbook objects, set each workbook's calculation mode to Manual via reflection, handle missing files and loading exceptions, and store the successfully loaded workbooks in a List<Workbook> for further processing.
class Program
{
    static void Main()
    {
        // Paths of the workbooks to be loaded
        var filePaths = new List<string>
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx",
            "Workbook3.xlsx"
        };

        // List that will hold the loaded workbooks
        var workbooks = new List<Workbook>();

        foreach (var path in filePaths)
        {
            try
            {
                // Verify that the file exists before attempting to load it
                if (!File.Exists(path))
                {
                    Console.WriteLine($"File not found: {path}");
                    continue;
                }

                // Load the workbook from the file
                var wb = new Workbook(path);

                // Attempt to set calculation mode to Manual via reflection (avoids compile‑time enum reference)
                try
                {
                    var settings = wb.Settings;
                    PropertyInfo calcModeProp = settings.GetType().GetProperty("CalcMode");
                    if (calcModeProp != null && calcModeProp.CanWrite)
                    {
                        // Parse the enum value "Manual" from the property type
                        object manualValue = Enum.Parse(calcModeProp.PropertyType, "Manual");
                        calcModeProp.SetValue(settings, manualValue);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unable to set manual calculation mode for '{path}': {ex.Message}");
                }

                // Add the workbook to the collection
                workbooks.Add(wb);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook '{path}': {ex.Message}");
            }
        }

        // At this point 'workbooks' contains all successfully loaded workbooks.
        Console.WriteLine($"Loaded {workbooks.Count} workbook(s) with manual calculation mode where possible.");
    }
}
