// Title: Sort Excel data with Aspose.Cells while keeping a merged header row intact (C#)
// AI Prompts: Sort a worksheet range by the first column using Aspose.Cells DataSorter while preserving a merged header row. | Configure DataSorter.HasHeaders and disable merge handling to sort data without affecting merged cells in C#. | Show how to define a CellArea and apply DataSorter.Sort so the merged header stays unchanged.
// Common Searches: C# Aspose.Cells sort range keep merged header row | How to prevent merged cells from moving when sorting with Aspose.Cells DataSorter | Aspose.Cells DataSorter disable merge handling example | Sorting Excel sheet with merged header using Aspose.Cells .NET | Preserve merged cells during sort Aspose.Cells C#
// Tags: Aspose.Cells DataSorter sort range with merged header | C# preserve merged cells during Excel sort | DataSorter.HasHeaders option example | CellArea definition for sorting Aspose.Cells | disable merge handling Aspose.Cells DataSorter

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, merges cells A1:C1 as a header, fills rows with fruit data, sets DataSorter.HasHeaders = true, defines a CellArea covering the header and data rows, sorts by the first column while keeping the merged header intact, and saves the workbook as SortedPreservingMergedCells.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // ------------------------------------------------------------
                // Prepare sample data with a merged header row
                // ------------------------------------------------------------

                // Merge cells A1:C1 to act as a header that should stay intact after sorting
                cells.Merge(0, 0, 1, 3);
                cells[0, 0].PutValue("Merged Header");

                // Add data below the header (rows 2‑5, columns A‑C)
                cells["A2"].PutValue("Banana");
                cells["B2"].PutValue(30);
                cells["C2"].PutValue("Yellow");

                cells["A3"].PutValue("Apple");
                cells["B3"].PutValue(10);
                cells["C3"].PutValue("Red");

                cells["A4"].PutValue("Cherry");
                cells["B4"].PutValue(20);
                cells["C4"].PutValue("Red");

                cells["A5"].PutValue("Date");
                cells["B5"].PutValue(40);
                cells["C5"].PutValue("Brown");

                // ------------------------------------------------------------
                // Configure the DataSorter
                // ------------------------------------------------------------

                // Get the DataSorter from the workbook
                DataSorter sorter = workbook.DataSorter;

                // The data has a header row (the merged cells). Setting HasHeaders = true
                // tells the sorter to keep the first row in place.
                sorter.HasHeaders = true;

                // Sort by the first column (Fruit name) in ascending order
                sorter.AddKey(0, SortOrder.Ascending);

                // ------------------------------------------------------------
                // Perform the sort
                // ------------------------------------------------------------

                // Define the area to sort: from row 0 (header) to row 4 (last data row),
                // and from column 0 to column 2 (A‑C). The header row is included but
                // will remain fixed because HasHeaders = true.
                CellArea sortArea = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 4,
                    EndColumn = 2
                };

                sorter.Sort(worksheet.Cells, sortArea);

                // ------------------------------------------------------------
                // Save the result
                // ------------------------------------------------------------

                string outputPath = "SortedPreservingMergedCells.xlsx";

                // Ensure the directory exists (use current directory if no path is provided)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
