// Title: Delete all empty rows and columns from the first worksheet of an Excel file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# delete empty rows and columns from the first sheet | How to remove blank rows and columns in an Excel workbook with Aspose.Cells | C# example for Cells.DeleteBlankRows and DeleteBlankColumns in Aspose.Cells | Save cleaned Excel file after deleting blank rows using Aspose.Cells | Check input file existence before processing with Aspose.Cells in C#
// Tags: Aspose.Cells delete blank rows C# | Aspose.Cells delete blank columns .NET | remove empty rows worksheet Aspose.Cells | clean Excel worksheet blank cells Aspose | Aspose.Cells workbook save after cleanup

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input Excel file, loads it with Aspose.Cells, accesses the first worksheet, removes all empty rows and columns using Cells.DeleteBlankRows and Cells.DeleteBlankColumns, then saves the cleaned workbook to a new file while handling possible exceptions.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Delete blank rows and columns in the worksheet
            sheet.Cells.DeleteBlankRows();
            sheet.Cells.DeleteBlankColumns();

            // Save the modified workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
