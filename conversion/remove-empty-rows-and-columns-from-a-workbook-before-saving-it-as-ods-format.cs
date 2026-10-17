// Title: Delete empty rows and columns from an Excel workbook and save as ODS using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells remove blank rows and columns before ODS export | How to clean Excel worksheets of empty rows/columns with Aspose.Cells .NET | Export a cleaned Excel file to ODS format using Aspose.Cells in C# | Delete completely empty rows and columns in each worksheet with Aspose.Cells | Aspose.Cells DeleteBlankRows DeleteBlankColumns example for ODS conversion
// Tags: Aspose.Cells DeleteBlankRows | Aspose.Cells DeleteBlankColumns | Aspose.Cells ODS export cleanup | C# remove empty rows columns Excel | Excel to ODS conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# program loads an existing Excel workbook, iterates through each worksheet, removes fully empty rows and columns using DeleteBlankRows and DeleteBlankColumns, and then saves the cleaned workbook as an ODS file, with basic file existence checks and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.ods";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Process each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Remove all completely empty rows
                sheet.Cells.DeleteBlankRows();

                // Remove all completely empty columns
                sheet.Cells.DeleteBlankColumns();
            }

            // Save the cleaned workbook in ODS format
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
