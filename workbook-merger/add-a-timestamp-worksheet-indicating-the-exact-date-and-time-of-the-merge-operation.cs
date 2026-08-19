// Title: Add a Merge Timestamp Worksheet While Merging Workbooks with Aspose.Cells for .NET (C#)
// Description: C# sample that validates two Excel files, loads them into Aspose.Cells workbooks, merges the source into the destination, creates a "MergeTimestamp" sheet, writes the current date‑time to cell A1 with a built‑in format, and saves the result as "MergedWithTimestamp.xlsx" with robust error handling.
// Keywords: Aspose.Cells merge workbooks C# | add timestamp worksheet Aspose.Cells | Excel merge audit trail .NET | record merge date time Aspose | C# combine Excel files with timestamp | date‑time format cell Aspose.Cells | version tracking Excel merge
// Common Searches: how to add a timestamp sheet after merging Excel files using Aspose.Cells | C# merge two workbooks and log merge date | Aspose.Cells combine workbooks and format date cell | add merge time to Excel workbook programmatically .NET | audit‑ready Excel merge with timestamp worksheet
// Developer Intent: Merge two Excel workbooks and automatically insert a worksheet that records the exact merge date and time.
// Use Cases: Create audit‑ready reports that show when data was consolidated. | Automate daily data aggregation with a built‑in version‑control sheet. | Generate backup copies of combined workbooks that include a timestamp for change management.
// AI Prompts: Generate C# code using Aspose.Cells to merge two Excel files and add a "MergeTimestamp" worksheet with the current date‑time in cell A1 formatted as mm/dd/yyyy hh:mm:ss. | Enhance the merge‑with‑timestamp example with checks for an existing timestamp sheet and allow a custom date format string. | Write a reusable method that accepts source and destination paths, merges the workbooks, adds a timestamp sheet, and returns the path of the saved file.

using System;
using System.IO;
using Aspose.Cells;

// C# sample that validates two Excel files, loads them into Aspose.Cells workbooks, merges the source into the destination, creates a "MergeTimestamp" sheet, writes the current date‑time to cell A1 with a built‑in format, and saves the result as "MergedWithTimestamp.xlsx" with robust error handling.
class MergeWithTimestamp
{
    static void Main()
    {
        try
        {
            // Verify that the source and destination files exist
            const string sourcePath = "Source.xlsx";
            const string destinationPath = "Destination.xlsx";

            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            if (!File.Exists(destinationPath))
                throw new FileNotFoundException($"Destination file not found: {destinationPath}");

            // Load the workbooks to be merged
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Workbook destinationWorkbook = new Workbook(destinationPath);

            // Combine the source workbook into the destination workbook
            destinationWorkbook.Combine(sourceWorkbook);

            // Add a new worksheet that will hold the merge timestamp
            Worksheet timestampSheet = destinationWorkbook.Worksheets.Add("MergeTimestamp");

            // Write the current date and time into cell A1
            Cell timestampCell = timestampSheet.Cells["A1"];
            timestampCell.PutValue(DateTime.Now);

            // Apply a standard date‑time number format (e.g., mm/dd/yyyy hh:mm:ss)
            Style dateStyle = destinationWorkbook.CreateStyle();
            dateStyle.Number = 22; // Built‑in date‑time format
            timestampCell.SetStyle(dateStyle);

            // Save the resulting workbook
            const string outputPath = "MergedWithTimestamp.xlsx";
            destinationWorkbook.Save(outputPath);
            Console.WriteLine($"Merged workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
