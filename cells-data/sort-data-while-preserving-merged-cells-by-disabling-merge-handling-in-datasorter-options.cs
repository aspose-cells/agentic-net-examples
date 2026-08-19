// Title: C# – Sort Excel data while preserving merged header cells with Aspose.Cells DataSorter
// Description: Demonstrates how to sort a worksheet range by a specific column using Aspose.Cells.DataSorter in .NET, with HasHeaders set to true so merged header cells remain unchanged, and saves the result to an XLSX file.
// Keywords: Aspose.Cells | DataSorter | C# sort merged cells | preserve merged header | HasHeaders true | Excel sort range .NET | disable merge handling | CellArea sort | Aspose.Cells example
// Common Searches: Aspose.Cells sort range without breaking merged cells | C# preserve merged header when sorting Excel with Aspose | DataSorter keep merged cells intact | How to disable merge handling in Aspose.Cells DataSorter | Sort Excel sheet by column while keeping merged header
// Developer Intent: Sort worksheet rows by a column while leaving merged header cells unchanged using Aspose.Cells in C#.
// Use Cases: Reorder a category/value list in a report where the title row spans multiple columns. | Apply an ascending numeric sort to a financial table without affecting a merged title row. | Organize product inventory data while preserving a merged block containing the report name and date. | Generate a sorted export of survey results that includes a merged header for the questionnaire title.
// AI Prompts: Write C# code that uses Aspose.Cells.DataSorter to sort a worksheet range by column B and keep merged header cells intact. | Explain how setting DataSorter.HasHeaders = true disables merge handling during sorting. | Show how to define a CellArea for sorting when the first row contains merged cells. | Provide a step‑by‑step guide to sort multiple columns in Aspose.Cells while preserving any merged cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to sort a worksheet range by a specific column using Aspose.Cells.DataSorter in .NET, with HasHeaders set to true so merged header cells remain unchanged, and saves the result to an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data (including a header row).
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("Fruit");
                cells["B2"].PutValue(30);
                cells["A3"].PutValue("Vegetable");
                cells["B3"].PutValue(20);
                cells["A4"].PutValue("Grain");
                cells["B4"].PutValue(10);

                // Merge the header cells to demonstrate that the merge is preserved after sorting.
                cells.Merge(0, 0, 1, 2); // Merge A1:B1

                // Configure the DataSorter.
                DataSorter sorter = workbook.DataSorter;
                sorter.HasHeaders = true;                 // First row is a header (merged cells)
                sorter.AddKey(1, SortOrder.Ascending);    // Sort by the second column (Value)

                // Define the sort area (including the merged header).
                CellArea sortArea = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 3,
                    EndColumn = 1
                };

                // Perform the sort. Merged cells are not altered because the sorter
                // does not process merge handling when HasHeaders is true.
                sorter.Sort(cells, sortArea);

                // Save the workbook.
                string outputPath = "SortedPreservingMergedCells.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
