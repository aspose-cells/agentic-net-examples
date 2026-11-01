// Title: Create an external workbook reference formula that auto‑updates when the source file moves using Aspose.Cells for .NET
// Common Searches: how to set an external workbook reference in Aspose.Cells C# that follows moved source file | Aspose.Cells C# external SUM formula with dynamic file path | update linked workbook formula path after moving source Excel file using Aspose.Cells | check external source file existence before assigning formula in Aspose.Cells C# | build external reference formula with folder and file variables in Aspose.Cells .NET
// Tags: Aspose.Cells external workbook linking | C# dynamic Excel formula path | verify external workbook existence Aspose.Cells | auto‑refresh linked workbook reference Aspose.Cells | construct external SUM formula Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, constructs an external SUM formula that references a source workbook using folder, file, and sheet variables, verifies the source file exists, assigns the formula to cell A1, and saves the result as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define parts of the external reference
            string externalFolder = @"C:\Data";
            string externalFile = "Source.xlsx";
            string externalSheet = "Sheet1";

            // Verify that the external source file exists to avoid FileNotFoundException
            string externalFullPath = Path.Combine(externalFolder, externalFile);
            if (!File.Exists(externalFullPath))
            {
                Console.WriteLine($"Warning: External source file not found: {externalFullPath}");
            }

            // Build the external formula (e.g., =SUM('C:\Data\[Source.xlsx]Sheet1'!A1:A10))
            string formula = $"=SUM('{externalFolder}\\[{externalFile}]{externalSheet}'!A1:A10)";

            // Assign the formula to a cell
            sheet.Cells["A1"].Formula = formula;

            // Save the workbook
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(resultPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
