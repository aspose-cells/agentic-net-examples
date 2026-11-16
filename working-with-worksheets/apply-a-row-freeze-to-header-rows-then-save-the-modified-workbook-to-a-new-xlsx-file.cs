// Title: Freeze the header row of an Excel worksheet and save as a new XLSX file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# freeze top row of worksheet | How to keep header row visible while scrolling with Aspose.Cells | Save workbook after applying FreezePanes in Aspose.Cells .NET | Freeze first row and export to a new file using Aspose.Cells C# | Set freeze panes for header row in an existing XLSX with Aspose.Cells
// Tags: Aspose.Cells FreezePanes header row | C# freeze first row Excel | Aspose.Cells save workbook as XLSX | Excel freeze panes using Aspose.Cells .NET | modify worksheet freeze rows C# | Aspose.Cells workbook freeze rows

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file with Aspose.Cells, freezes the first (header) row using the FreezePanes method, and saves the updated workbook to a new XLSX file, handling file existence and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify the required sheet)
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze the first row (header row)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns
            // Here we freeze rows above row index 1 (i.e., row 1) and no columns.
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the modified workbook to a new XLSX file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
