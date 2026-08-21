// Title: C# – Detect Cells with Leading Apostrophe and Rich Text Using Aspose.Cells
// Description: Loads a workbook, scans the used range, checks each non‑empty cell for the QuotePrefix flag (leading apostrophe) and the IsRichText() property, logs the cell address when both are true, and optionally saves the unchanged file.
// Keywords: Aspose.Cells leading apostrophe detection | C# find rich text cells | QuotePrefix Aspose.Cells | IsRichText method | scan used range Aspose | log cell addresses .NET | detect apostrophe rich text | Aspose.Cells data validation | Excel apostrophe detection C# | rich text cell audit
// Common Searches: How to find cells with a leading apostrophe and rich text in Aspose.Cells C# | Aspose.Cells QuotePrefix and IsRichText example | Detect apostrophe‑prefixed rich‑text cells using .NET | Log addresses of cells that have both QuotePrefix and rich text | C# code to audit Excel for leading apostrophes with formatting
// Developer Intent: Log addresses of cells that contain both a leading apostrophe and rich‑text formatting.
// Use Cases: Validate imported spreadsheet data for unintended apostrophe prefixes while preserving rich‑text formatting. | Generate an audit report of cells that combine quote prefix and rich‑text for cleanup before publishing. | Perform an automated quality check to ensure hidden apostrophes do not affect calculations. | Identify cells that need special handling during custom export or transformation processes.
// AI Prompts: Create a C# function with Aspose.Cells that returns a List<string> of cell names where QuotePrefix is true and IsRichText() is true. | Rewrite the nested loops to iterate only over non‑empty cells using Cells.FindAll and log matching addresses. | Show how to remove the leading apostrophe from a cell without losing its rich‑text formatting. | Explain performance considerations when scanning large worksheets for QuotePrefix and rich‑text cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, scans the used range, checks each non‑empty cell for the QuotePrefix flag (leading apostrophe) and the IsRichText() property, logs the cell address when both are true, and optionally saves the unchanged file.
    public class DetectApostropheRichText
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Determine the used range of the worksheet
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Skip empty cells
                        if (cell.Type == CellValueType.IsNull)
                            continue;

                        // Leading apostrophe is indicated by the QuotePrefix style flag
                        bool hasLeadingApostrophe = cell.GetStyle().QuotePrefix;

                        // Rich‑text is indicated by the IsRichText method
                        bool isRichText = cell.IsRichText();

                        // Log cells that satisfy both conditions
                        if (hasLeadingApostrophe && isRichText)
                        {
                            Console.WriteLine($"Cell {cell.Name} contains a leading apostrophe and rich text.");
                        }
                    }
                }

                // Optionally save the workbook (unchanged in this example)
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
