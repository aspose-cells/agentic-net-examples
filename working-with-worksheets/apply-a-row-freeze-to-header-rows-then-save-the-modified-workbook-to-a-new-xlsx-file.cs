// Title: Freeze the header row of an existing XLSX workbook using Aspose.Cells for .NET and save it as a new file
// AI Prompts: Write C# code with Aspose.Cells to freeze the first worksheet row and save the workbook to a different XLSX file. | Create a .NET program that loads an Excel file, applies FreezePanes to the top row, and writes the result to a new workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# freeze top row and save as new workbook | How to apply FreezePanes to header row in a .NET Excel file | C# program to freeze first row of an existing XLSX using Aspose.Cells
// Tags: Aspose.Cells FreezePanes header row | C# freeze first row Excel workbook | save modified workbook as new XLSX Aspose.Cells | load existing XLSX apply row freeze .NET | worksheet FreezePanes example C#

using System;
using System.IO;
using Aspose.Cells;

// Loads input.xlsx, freezes the top row of the first worksheet with FreezePanes, and saves the updated workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify the desired sheet index/name)
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze the top row (header). FreezePanes(row, column, totalRows, totalColumns)
            // row = 1 (rows above 1 are frozen), column = 0 (no columns frozen),
            // totalRows = 1 (freeze one row), totalColumns = 0 (no columns frozen)
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the modified workbook to a new XLSX file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
