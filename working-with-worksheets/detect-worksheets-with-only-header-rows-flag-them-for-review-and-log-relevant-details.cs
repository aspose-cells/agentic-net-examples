// Title: C# – Detect Header‑Only Worksheets and Flag Them with a Custom Property using Aspose.Cells
// Description: Loads a workbook, scans each worksheet, and identifies sheets where only the first row contains data. When such a sheet is found, a custom property "NeedsReview" is added and the worksheet name, header row index, and column count are logged before saving the file.
// Keywords: Aspose.Cells header detection C# | flag worksheet needs review | custom property Aspose.Cells | MaxDataRow MaxDataColumn Excel .NET | identify empty data rows Excel | C# Excel worksheet validation | log worksheet details Aspose.Cells
// Common Searches: How to find worksheets that only have a header row with Aspose.Cells | Add a custom property to mark Excel sheets for review in .NET | Log name and column count of header‑only worksheets using C# | Detect empty data rows in Excel workbooks with Aspose.Cells
// Developer Intent: Find worksheets that contain only a header row, mark them for review, and output their key details.
// Use Cases: Automatically skip or flag sheets that lack data before running analytics pipelines. | Create a pre‑publish audit report listing all header‑only worksheets in a workbook. | Route flagged sheets to a data‑quality team via the "NeedsReview" property.
// AI Prompts: Generate C# code with Aspose.Cells that scans every worksheet, detects if only row 0 has values, adds a custom property "NeedsReview" set to true, and prints the sheet name and column count. | Show how to use MaxDataRow and MaxDataColumn to determine whether a worksheet consists solely of a header row in Aspose.Cells for .NET. | Explain a workflow that flags header‑only worksheets with a custom property and logs their details for downstream processing.

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetHeaderDetection
{
    // Loads a workbook, scans each worksheet, and identifies sheets where only the first row contains data. When such a sheet is found, a custom property "NeedsReview" is added and the worksheet name, header row index, and column count are logged before saving the file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the last row that contains any data (zero‑based index)
                    int lastDataRow = sheet.Cells.MaxDataRow;

                    // Check if the first row has any non‑empty cells
                    bool firstRowHasData = false;
                    for (int col = 0; col <= sheet.Cells.MaxDataColumn; col++)
                    {
                        if (!string.IsNullOrEmpty(sheet.Cells[0, col].StringValue))
                        {
                            firstRowHasData = true;
                            break;
                        }
                    }

                    // If only the first row contains data, treat it as a header‑only sheet
                    if (firstRowHasData && lastDataRow == 0)
                    {
                        // Flag the worksheet for review using a custom property (value stored as string)
                        sheet.CustomProperties.Add("NeedsReview", "true");

                        // Log relevant details
                        Console.WriteLine($"Worksheet \"{sheet.Name}\" contains only header rows.");
                        Console.WriteLine($"  Header row index: 0");
                        Console.WriteLine($"  Number of columns with data: {sheet.Cells.MaxDataColumn + 1}");
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
