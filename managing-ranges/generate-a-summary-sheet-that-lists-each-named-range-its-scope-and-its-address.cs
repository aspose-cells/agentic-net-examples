using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;   // Alias to avoid conflict with System.Range

namespace AsposeCellsNamedRangeSummary
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample data and named ranges to demonstrate the summary
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";
                sheet1.Cells["A1"].PutValue("Item");
                sheet1.Cells["A2"].PutValue("Apple");
                sheet1.Cells["A3"].PutValue("Banana");

                // Create a workbook‑scoped named range
                int globalIndex = workbook.Worksheets.Names.Add("GlobalRange");
                workbook.Worksheets.Names[globalIndex].RefersTo = "=Sheet1!$A$2:$A$3";

                // Create a worksheet‑scoped named range on Sheet1
                int sheetIndex = workbook.Worksheets.Names.Add("Sheet1!LocalRange");
                Name localName = workbook.Worksheets.Names[sheetIndex];
                localName.RefersTo = "=Sheet1!$A$1:$A$3";
                localName.SheetIndex = 0; // 0 = global, otherwise one‑based sheet index

                // Add a second worksheet with its own named range
                Worksheet sheet2 = workbook.Worksheets.Add("Data");
                sheet2.Cells["B1"].PutValue(10);
                sheet2.Cells["B2"].PutValue(20);
                int sheet2Index = workbook.Worksheets.Names.Add("DataRange");
                workbook.Worksheets.Names[sheet2Index].RefersTo = "=Data!$B$1:$B$2";

                // Create a summary worksheet
                Worksheet summary = workbook.Worksheets.Add("Summary");
                // Write header row
                summary.Cells["A1"].PutValue("Name");
                summary.Cells["B1"].PutValue("Scope");
                summary.Cells["C1"].PutValue("Address");

                // Iterate over all defined names
                NameCollection names = workbook.Worksheets.Names;
                int row = 1; // zero‑based index; row 1 is the second row (after header)

                foreach (Name name in names)
                {
                    // Determine scope: Workbook (global) or specific worksheet name
                    string scope;
                    if (name.SheetIndex == 0)
                    {
                        scope = "Workbook";
                    }
                    else
                    {
                        // SheetIndex is one‑based; convert to zero‑based for collection access
                        int wsIndex = name.SheetIndex - 1;
                        if (wsIndex >= 0 && wsIndex < workbook.Worksheets.Count)
                            scope = workbook.Worksheets[wsIndex].Name;
                        else
                            scope = "Unknown";
                    }

                    // Get the address of the range; use GetRange() if possible
                    string address = string.Empty;
                    try
                    {
                        AsposeRange rng = name.GetRange();
                        if (rng != null)
                            address = rng.Address;
                    }
                    catch
                    {
                        // Fallback to RefersTo string (remove leading '=')
                        address = name.RefersTo?.TrimStart('=') ?? string.Empty;
                    }

                    // Write details to the summary sheet
                    summary.Cells[row, 0].PutValue(name.Text);   // Column A
                    summary.Cells[row, 1].PutValue(scope);      // Column B
                    summary.Cells[row, 2].PutValue(address);    // Column C

                    row++;
                }

                // Save the workbook
                string outputPath = "NamedRangeSummary.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}