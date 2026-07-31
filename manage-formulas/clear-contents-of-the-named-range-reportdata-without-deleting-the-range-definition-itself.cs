// Title: C# – Clear the contents of the named range "ReportData" while keeping its definition (Aspose.Cells for .NET)
// Description: Demonstrates how to create or load a workbook, define a named range called ReportData, obtain its Aspose.Range object, call ClearContents() to erase all cell values without removing the name, and save the file.
// Keywords: Aspose.Cells C# clear named range | ClearContents Aspose.Range | preserve named range definition | remove values from named range .NET | Aspose.Cells example GitHub | C# workbook named range manipulation | Aspose.Cells clear cells without deleting name
// Common Searches: Aspose.Cells clear values of a named range | How to keep a named range after clearing its cells in .NET | ClearContents vs Clear in Aspose.Range | C# code to empty a named range but retain the name | Aspose.Cells example for clearing ReportData range
// Developer Intent: Erase all data inside the named range "ReportData" while leaving the range name and its reference unchanged.
// Use Cases: Refresh a report template by wiping old results before writing new data. | Reset input sections of a workbook without breaking formulas that rely on the named range. | Prepare a workbook for reuse in batch processing by clearing calculation outputs while preserving named range links.
// AI Prompts: Show C# code that clears the contents of a named range in Aspose.Cells without deleting the name. | Explain when to use ClearContents versus Clear on an Aspose.Range object. | Give an example of repopulating a named range after calling ClearContents in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to create or load a workbook, define a named range called ReportData, obtain its Aspose.Range object, call ClearContents() to erase all cell values without removing the name, and save the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (replace with Workbook wb = new Workbook("input.xlsx"); to load an existing file)
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Populate some data in the area that will be named "ReportData"
            ws.Cells["A1"].PutValue("Sample 1");
            ws.Cells["B2"].PutValue(42);

            // Retrieve the named range "ReportData" if it exists; otherwise create it
            Name reportName = wb.Worksheets.Names["ReportData"];
            if (reportName == null)
            {
                // Add returns the index of the newly created name
                int nameIndex = wb.Worksheets.Names.Add("ReportData");
                reportName = wb.Worksheets.Names[nameIndex];
            }

            // Set the reference of the named range to the desired area
            reportName.RefersTo = ws.Name + "!$A$1:$B$2";

            // Get the Range object associated with the named range
            AsposeRange reportRange = reportName.GetRange();

            // Clear only the contents of the range, preserving the range definition
            reportRange.ClearContents();

            // Save the workbook (replace with desired path)
            string outputPath = "ClearedReportData.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
