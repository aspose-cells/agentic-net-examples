// Title: Generate an Excel data‑type summary sheet by counting CellValueType occurrences with Aspose.Cells for .NET (C#)
// Common Searches: how to count different CellValueType values in an Excel worksheet using Aspose.Cells C# | Aspose.Cells C# example to create a summary sheet with data type frequencies | enumerate used range cells and get type counts in Aspose.Cells .NET | C# code to produce a data type report for an XLSX file with Aspose.Cells | auto‑fit columns after writing a summary report with Aspose.Cells in C#
// Tags: Aspose.Cells enumerate cell value types | Aspose.Cells create data type summary worksheet | C# count Excel CellValueType occurrences | Aspose.Cells auto-fit columns after report generation | Aspose.Cells dictionary initialization for CellValueType

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace DataTypeSummaryReport
{
    // The sample loads an existing XLSX workbook (or creates one with sample data), determines the used range of the first worksheet, iterates through each cell, counts how many times each CellValueType appears using a dictionary, adds a new "Summary" worksheet, writes non‑zero type counts as a two‑column table, auto‑fits the columns, and saves the workbook as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one with sample data
                if (File.Exists(inputPath))
                {
                    // LoadOptions can be used to specify password if needed; here we assume none
                    var loadOptions = new LoadOptions(LoadFormat.Xlsx);
                    workbook = new Workbook(inputPath, loadOptions);
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Name = "Sheet1";
                    ws.Cells["A1"].PutValue("Sample Text");
                    ws.Cells["B2"].PutValue(12345);
                    ws.Cells["C3"].PutValue(DateTime.Now);
                }

                // Get the first worksheet (or specify the sheet you want to analyze)
                Worksheet sourceSheet = workbook.Worksheets[0];

                // Dictionary to hold counts for each cell value type
                Dictionary<CellValueType, int> typeCounts = new Dictionary<CellValueType, int>();

                // Initialize counts for all possible types
                foreach (CellValueType type in Enum.GetValues(typeof(CellValueType)))
                {
                    typeCounts[type] = 0;
                }

                // Determine the used range of the worksheet
                int maxRow = sourceSheet.Cells.MaxDataRow;
                int maxColumn = sourceSheet.Cells.MaxDataColumn;

                // Enumerate all cells within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = sourceSheet.Cells[row, col];
                        CellValueType cellType = cell.Type;

                        // Increment the corresponding counter
                        if (typeCounts.ContainsKey(cellType))
                        {
                            typeCounts[cellType]++;
                        }
                        else
                        {
                            // In case a new type appears (unlikely), add it to the dictionary
                            typeCounts[cellType] = 1;
                        }
                    }
                }

                // Create a new worksheet for the summary report
                Worksheet summarySheet = workbook.Worksheets.Add("Summary");

                // Write header titles
                summarySheet.Cells[0, 0].PutValue("Cell Value Type");
                summarySheet.Cells[0, 1].PutValue("Count");

                // Populate the summary data
                int summaryRow = 1;
                foreach (var kvp in typeCounts)
                {
                    // Skip types with zero count to keep the report concise
                    if (kvp.Value == 0) continue;

                    summarySheet.Cells[summaryRow, 0].PutValue(kvp.Key.ToString());
                    summarySheet.Cells[summaryRow, 1].PutValue(kvp.Value);
                    summaryRow++;
                }

                // Auto-fit columns for better readability
                summarySheet.AutoFitColumns();

                // Save the workbook with the summary sheet
                workbook.Save(outputPath);
                Console.WriteLine($"Summary report saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
