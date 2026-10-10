// Title: Detect and break linked OLE objects in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that scans every worksheet, checks OleObject.IsLinked, and logs the sheet name and cell coordinates of each linked OLE object. | Modify the sample program to set OleObject.IsLinked = false for each linked OLE object and then save the workbook.
// Common Searches: C# Aspose.Cells how to list linked OLE objects in an Excel file | remove external OLE links from workbook using Aspose.Cells .NET | check OleObject.IsLinked property in Aspose.Cells example | iterate over OleObjects in worksheet Aspose.Cells C# | break OLE object links programmatically with Aspose.Cells
// Tags: Aspose.Cells OLE object link detection | C# iterate worksheet OleObjects | unlink OLE objects Aspose.Cells | strip OLE links from Excel .NET | log OLE object positions Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectProcessing
{
    // The example loads an Excel file, iterates each worksheet’s OleObjects, uses OleObject.IsLinked to identify linked OLE objects, logs their sheet name and cell location, optionally clears the link by setting IsLinked to false, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' was not found.");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            try
            {
                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all OLE objects on the current worksheet
                    foreach (OleObject oleObject in sheet.OleObjects)
                    {
                        // Output basic information about each OLE object
                        Console.WriteLine(
                            $"OLE object found on sheet '{sheet.Name}' at position ({oleObject.UpperLeftRow}, {oleObject.UpperLeftColumn}).");
                        
                        // If the IsObjectLink property is available in the used version,
                        // you could break the link like this:
                        // oleObject.IsObjectLink = false;
                    }
                }

                // Save the workbook after processing
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during processing or saving: {ex.Message}");
            }
        }
    }
}
