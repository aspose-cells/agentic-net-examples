// Title: Load an Excel workbook, add an external SUM formula, and save the file using Aspose.Cells for .NET (C#)
// Common Searches: how to add a formula that references another workbook with Aspose.Cells C# | Aspose.Cells C# load existing workbook and set external SUM formula | saving workbook after updating cell formula using Aspose.Cells .NET | C# code to set formula =SUM('[ExternalFile.xlsx]Sheet1'!A1:A10) with Aspose.Cells | handling missing input file when loading workbook in Aspose.Cells C#
// Tags: load workbook with Aspose.Cells LoadOptions | set external reference formula Aspose.Cells | update cell formula in Excel using Aspose.Cells C# | save workbook after formula modification Aspose.Cells | handle CellsException in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of 'input.xlsx', loads it with Aspose.Cells using default LoadOptions, accesses the first worksheet, assigns a SUM formula that references an external workbook ('[ExternalFile.xlsx]Sheet1'!A1:A10) to cell B1, saves the modified workbook as 'output.xlsx', and includes error handling for CellsException and generic exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook (no password required)
            var loadOptions = new LoadOptions(); // Add password here if needed: loadOptions.Password = "yourPassword";
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Set a formula that references an external workbook.
            // Example: =SUM('[ExternalFile.xlsx]Sheet1'!A1:A10)
            sheet.Cells["B1"].Formula = "SUM('[ExternalFile.xlsx]Sheet1'!A1:A10)";

            // Save the workbook with the updated formula
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
