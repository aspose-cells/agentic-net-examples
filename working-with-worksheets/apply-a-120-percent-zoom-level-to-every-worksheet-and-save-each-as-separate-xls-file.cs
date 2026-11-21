// Title: Set 120% zoom for each worksheet and export each as a separate XLS file with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set worksheet zoom to 120 percent and save as XLS | How to export each sheet of an Excel workbook to individual .xls files using Aspose.Cells | C# loop through worksheets, apply zoom, and create separate workbooks with Aspose.Cells | Copy a single worksheet to a new workbook and save in Excel97 format using Aspose.Cells | Batch process worksheets to change view zoom and export to legacy XLS in .NET
// Tags: worksheet zoom setting Aspose.Cells | export single worksheet to XLS Aspose.Cells | copy worksheet to new workbook C# | save workbook as Excel97 format Aspose.Cells | batch worksheet processing Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The program loads SourceWorkbook.xlsx, iterates through all worksheets, sets each sheet's Zoom to 120%, creates a new workbook containing only that sheet, and saves it as a separate .xls file named after the worksheet.
class Program
{
    static void Main()
    {
        try
        {
            // Define the source workbook path
            string sourcePath = "SourceWorkbook.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Iterate through all worksheets in the source workbook
            for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
            {
                Worksheet sheet = sourceWorkbook.Worksheets[i];

                // Apply a 120% zoom level to the current worksheet
                sheet.Zoom = 120;

                // Create a new workbook that will contain only this worksheet
                Workbook singleSheetWorkbook = new Workbook();

                // Remove the default empty sheet created with the new workbook
                singleSheetWorkbook.Worksheets.Clear();

                // Copy the current worksheet into the new workbook (use sheet name)
                singleSheetWorkbook.Worksheets.AddCopy(sheet.Name);

                // Define the output file name (one file per worksheet)
                string outputFileName = $"{sheet.Name}.xls";

                // Save the new workbook as an XLS file
                singleSheetWorkbook.Save(outputFileName, SaveFormat.Excel97To2003);
                Console.WriteLine($"Saved worksheet '{sheet.Name}' to '{outputFileName}'.");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
