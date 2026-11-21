// Title: Remove a slicer by its name from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: aspocells c# delete slicer by name from xlsx | how to programmatically remove a specific slicer in Excel with Aspose.Cells | c# reflection to access and delete slicers in an Excel file using Aspose.Cells | clean up unused slicers in a workbook with Aspose.Cells .NET | example code to remove a slicer from a worksheet using Aspose.Cells
// Tags: aspocells delete slicer | c# aspocells slicer removal | excel slicer cleanup aspocells | reflection based slicer access aspocells | remove unused slicer .net

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Tables; // Retained for compatibility if the API is available

// Loads an input.xlsx workbook, uses reflection to enumerate each worksheet's Slicers collection, finds the slicer whose Name matches the specified value (e.g., "MySlicer"), removes it via the collection's RemoveAt method, and saves the modified workbook to output.xlsx, with file existence checks and exception handling.
class DeleteSlicerExample
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Target slicer name
        string slicerName = "MySlicer";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            bool slicerRemoved = false;

            // Iterate through worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Use reflection to access the Slicers collection (avoids compile‑time dependency)
                PropertyInfo slicersProp = sheet.GetType().GetProperty("Slicers", BindingFlags.Public | BindingFlags.Instance);
                if (slicersProp == null) continue; // No slicer support in this worksheet

                object slicersObj = slicersProp.GetValue(sheet);
                if (slicersObj == null) continue;

                // Retrieve collection members via reflection
                PropertyInfo countProp = slicersObj.GetType().GetProperty("Count");
                MethodInfo getItemMethod = slicersObj.GetType().GetMethod("get_Item");
                MethodInfo removeAtMethod = slicersObj.GetType().GetMethod("RemoveAt");

                if (countProp == null || getItemMethod == null || removeAtMethod == null) continue;

                int count = (int)countProp.GetValue(slicersObj);

                // Loop backwards to safely remove items
                for (int i = count - 1; i >= 0; i--)
                {
                    object slicer = getItemMethod.Invoke(slicersObj, new object[] { i });
                    if (slicer == null) continue;

                    PropertyInfo nameProp = slicer.GetType().GetProperty("Name");
                    if (nameProp == null) continue;

                    string name = nameProp.GetValue(slicer) as string;
                    if (string.Equals(name, slicerName, StringComparison.OrdinalIgnoreCase))
                    {
                        // Remove the slicer
                        removeAtMethod.Invoke(slicersObj, new object[] { i });
                        slicerRemoved = true;
                        break; // Assuming unique slicer names
                    }
                }

                if (slicerRemoved) break;
            }

            if (!slicerRemoved)
            {
                Console.WriteLine($"Slicer \"{slicerName}\" not found in the workbook.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Graceful error handling
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
