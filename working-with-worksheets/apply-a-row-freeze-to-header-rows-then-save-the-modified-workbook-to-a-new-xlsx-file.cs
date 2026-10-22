// Title: Freeze the first row of an Excel worksheet and save as a new XLSX using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing XLSX file with Aspose.Cells, applies a freeze to the first row of the first worksheet, and writes the workbook to a new file. | Show how to call the FreezePanes method in Aspose.Cells to lock the header row and then save the updated workbook as a separate XLSX document.
// Common Searches: asp.net freeze first worksheet row using Aspose.Cells | c# example of FreezePanes for header row in Excel file | how to keep top row static in Excel with Aspose.Cells library | save workbook after freezing panes with Aspose.Cells C# | Aspose.Cells freeze panes on first sheet and export new XLSX
// Tags: Aspose.Cells FreezePanes header row | C# freeze top row Excel worksheet | Aspose.Cells save workbook as XLSX | programmatic row freeze Excel .NET | Excel row freeze using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing XLSX workbook, freezes the top header row of the first worksheet using Aspose.Cells' FreezePanes method, and saves the modified workbook to a new XLSX file.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook from the file
            var workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust the index if needed)
            var worksheet = workbook.Worksheets[0];

            // Freeze the top row (header row)
            // Parameters: row (first unfrozen row index), column (first unfrozen column index),
            // totalRows (number of rows to freeze), totalColumns (number of columns to freeze)
            worksheet.FreezePanes(1, 0, 1, 0);

            // Save the modified workbook to a new XLSX file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
