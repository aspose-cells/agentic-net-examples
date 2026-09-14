// Title: Set StandardWidth and auto‑fit a specific column range in Aspose.Cells for .NET (C#)
// AI Prompts: Configure the worksheet's StandardWidth, apply custom widths to columns A and B, then auto‑fit columns A‑C and read the resulting widths using Aspose.Cells in C#. | Show how to compare column width values before and after calling AutoFitColumns when some columns have been manually sized with SetColumnWidth.
// Common Searches: Aspose.Cells C# set default column width and then auto‑fit selected columns | How to use AutoFitColumns on a range after overriding column widths with SetColumnWidth in Aspose.Cells | Retrieve column width values before and after AutoFitColumns in Aspose.Cells .NET | C# example of StandardWidth property impact on auto‑fit behavior in an Excel workbook | Save workbook after adjusting column widths with StandardWidth and AutoFitColumns using Aspose.Cells
// Tags: Aspose.Cells StandardWidth property | Aspose.Cells AutoFitColumns range | Aspose.Cells SetColumnWidth override | C# Excel column width management | Aspose.Cells workbook save xlsx | Aspose.Cells column width before after AutoFit

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, sets the worksheet's StandardWidth to 20 characters, manually overrides widths for columns A and B, populates cells with varying text lengths, auto‑fits columns A‑C, prints column widths before and after auto‑fit, and saves the workbook as an XLSX file.
    public class StandardWidthAndAutoFitDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Set the default column width (in character units) for the worksheet
                cells.StandardWidth = 20.0;
                Console.WriteLine($"StandardWidth set to: {cells.StandardWidth}");

                // Manually override widths for the first two columns
                cells.SetColumnWidth(0, 10.0); // Column A narrower than standard
                cells.SetColumnWidth(1, 30.0); // Column B wider than standard
                Console.WriteLine($"Column A width before AutoFit: {cells.GetColumnWidth(0)}");
                Console.WriteLine($"Column B width before AutoFit: {cells.GetColumnWidth(1)}");
                Console.WriteLine($"Column C width before AutoFit (default): {cells.GetColumnWidth(2)}");

                // Populate cells with data that requires wider columns
                worksheet.Cells["A1"].PutValue("Short");
                worksheet.Cells["A2"].PutValue("A bit longer text");
                worksheet.Cells["B1"].PutValue("This is a very long piece of text that should trigger auto‑fit");
                worksheet.Cells["B2"].PutValue("Another long text entry for column B");
                worksheet.Cells["C1"].PutValue("Medium length");
                worksheet.Cells["C2"].PutValue("Extremely long text that will cause column C to expand when auto‑fit is applied");

                // Auto‑fit columns 0 through 2 (A, B, C)
                worksheet.AutoFitColumns(0, 2);

                // Display column widths after auto‑fit to observe which settings were overridden
                Console.WriteLine($"Column A width after AutoFit: {cells.GetColumnWidth(0)}");
                Console.WriteLine($"Column B width after AutoFit: {cells.GetColumnWidth(1)}");
                Console.WriteLine($"Column C width after AutoFit: {cells.GetColumnWidth(2)}");

                // Determine output file path and ensure the directory exists
                string outputFile = "StandardWidthAndAutoFitDemo.xlsx";
                string outputPath = Path.GetFullPath(outputFile);
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            StandardWidthAndAutoFitDemo.Run();
        }
    }
}
