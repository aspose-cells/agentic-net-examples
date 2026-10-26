// Title: Combine all worksheets from an XLSX workbook and save the merged workbook as an XLS (Excel 97‑2003) file using Aspose.Cells for .NET
// Common Searches: c# Aspose.Cells copy multiple sheets into a new workbook and export to Excel 97-2003 | how to programmatically merge worksheets from an .xlsx and generate a .xls file with Aspose.Cells | save merged workbook as XLS format using Aspose.Cells .NET API | convert combined XLSX sheets to legacy XLS using Aspose.Cells in C# | Aspose.Cells example for merging worksheets and saving as Excel 97-2003
// Tags: copy worksheets between workbooks Aspose.Cells | save workbook as Excel97To2003 format C# | merge multiple sheets and export to XLS Aspose | legacy XLS output from XLSX Aspose.Cells | combined workbook generation Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The code loads an XLSX file, creates a new workbook, copies each worksheet from the source into the new workbook, and then saves the merged workbook as an XLS file using the Excel97To2003 save format, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xls";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(inputPath);

            // Create a new workbook for the merged content
            Workbook mergedWorkbook = new Workbook();
            mergedWorkbook.Worksheets.Clear(); // Remove the default empty sheet

            // Copy each worksheet from the source workbook into the merged workbook
            foreach (Worksheet sheet in sourceWorkbook.Worksheets)
            {
                // Add a copy of the current sheet by its name
                mergedWorkbook.Worksheets.AddCopy(sheet.Name);
            }

            // Save the merged workbook in Excel 97-2003 format
            mergedWorkbook.Save(outputPath, SaveFormat.Excel97To2003);
            Console.WriteLine($"Merged workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
