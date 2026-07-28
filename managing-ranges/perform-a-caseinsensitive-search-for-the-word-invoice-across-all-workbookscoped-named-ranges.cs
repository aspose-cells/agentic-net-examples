// Title: Case‑insensitive search for "invoice" in all workbook‑scoped named ranges (Aspose.Cells for .NET)
// Description: Loads an Excel workbook, filters to workbook‑level named ranges, iterates each range, and uses FindOptions (case‑insensitive, contains) to locate cells that contain the word "invoice". Matching range name, cell address, and worksheet are written to the console, and the workbook is saved unchanged.
// Keywords: Aspose.Cells | C# | .NET | named ranges | workbook‑scoped names | global named ranges | case insensitive search | FindOptions | search text in Excel | invoice lookup | Excel automation example | GitHub code sample
// Common Searches: search text in all global named ranges Aspose.Cells | case insensitive find in workbook scoped names C# | Aspose.Cells find "invoice" in named ranges | C# code to locate cells containing a word in Excel named ranges | how to filter workbook‑level names with Aspose.Cells
// Developer Intent: Identify every cell that contains the word "invoice" within workbook‑scoped named ranges.
// Use Cases: Create an audit list of invoice references across multiple sheets for data validation. | Generate a report that logs worksheet, named range, and cell address for each occurrence of "invoice". | Trigger custom actions—such as applying formatting or raising alerts—when the term appears in any global named range.
// AI Prompts: Give a minimal Aspose.Cells C# snippet that searches case‑insensitively for a string in all workbook‑scoped named ranges and returns matching cell addresses. | Show how to modify the example to replace every found "invoice" with "receipt" across the named ranges. | Explain how to export the found cell details (worksheet, range name, address) to a CSV file instead of printing to the console.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsInvoiceSearch
{
    // Loads an Excel workbook, filters to workbook‑level named ranges, iterates each range, and uses FindOptions (case‑insensitive, contains) to locate cells that contain the word "invoice". Matching range name, cell address, and worksheet are written to the console, and the workbook is saved unchanged.
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
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure find options: case‑insensitive, search in values, match if cell contains the word
                FindOptions findOptions = new FindOptions
                {
                    CaseSensitive = false,
                    LookInType = LookInType.Values,
                    LookAtType = LookAtType.Contains
                };

                // Retrieve only workbook‑scoped (global) named ranges
                Name[] workbookNames = workbook.Worksheets.Names.Filter(NameScopeType.Workbook, -1);

                foreach (Name name in workbookNames)
                {
                    // Get the ranges referenced by the name
                    AsposeRange[] ranges = name.GetRanges();

                    if (ranges == null) continue; // skip if the name does not refer to a range

                    foreach (AsposeRange range in ranges)
                    {
                        // Define the search area based on the range dimensions
                        CellArea area = new CellArea
                        {
                            StartRow = range.FirstRow,
                            StartColumn = range.FirstColumn,
                            EndRow = range.FirstRow + range.RowCount - 1,
                            EndColumn = range.FirstColumn + range.ColumnCount - 1
                        };
                        findOptions.SetRange(area);

                        // Perform the search within the defined area
                        Worksheet ws = range.Worksheet;
                        Cell foundCell = ws.Cells.Find("invoice", null, findOptions);

                        if (foundCell != null)
                        {
                            Console.WriteLine(
                                $"Found \"invoice\" in named range \"{name.Text}\" at cell {foundCell.Name} (Worksheet: {ws.Name})");
                        }
                    }
                }

                // Save the workbook (optional – here we just save without modifications)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
