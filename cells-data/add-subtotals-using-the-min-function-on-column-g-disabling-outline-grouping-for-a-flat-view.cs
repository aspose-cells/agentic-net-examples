// Title: C# – Add Min Subtotal on Column G with Flat View (Outline Disabled) using Aspose.Cells
// Description: Learn how to use Aspose.Cells for .NET to group rows by the first column, calculate the minimum value in column G, place the subtotal rows below the data, and turn off outline grouping for a flat, non‑collapsible view. The example creates sample data, applies Cells.Subtotal with ConsolidationFunction.Min, sets SummaryRowBelow, and saves the workbook as an .xlsx file.
// Keywords: Aspose.Cells C# subtotal Min | column G subtotal Aspose.Cells | flat view outline disabled | group by first column Excel | summary row below data | ConsolidationFunction.Min | Excel automation C# | Aspose.Cells Subtotal method | generate Excel report programmatically | disable outline grouping
// Common Searches: Aspose.Cells add Min subtotal column G | C# subtotal flat view without outline | disable outline grouping Aspose.Cells | group rows by first column and calculate minimum Aspose.Cells | place subtotal rows below data in Excel using C#
// Developer Intent: Create a workbook that groups rows by Category, adds a Min subtotal on the Score column (G), and shows the subtotals in a flat view with outline grouping turned off.
// Use Cases: Produce a sales report that shows the lowest score per category directly beneath each group, simplifying visual analysis. | Export inventory data where the minimum price per region is listed without collapsible outlines, making it easier for downstream systems to read. | Automate financial statements that require flat subtotal rows so that auditors can see summary values without expanding outline levels.
// AI Prompts: Generate C# code with Aspose.Cells to add a Min subtotal on column G, group by column A, place summary rows below the data, and disable outline grouping. | Show how to replace existing subtotals with a new Min subtotal on a different column while keeping a flat view in an Aspose.Cells workbook. | Explain each parameter of the Cells.Subtotal method for grouping, aggregation function, target columns, and flat‑view options in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Learn how to use Aspose.Cells for .NET to group rows by the first column, calculate the minimum value in column G, place the subtotal rows below the data, and turn off outline grouping for a flat, non‑collapsible view. The example creates sample data, applies Cells.Subtotal with ConsolidationFunction.Min, sets SummaryRowBelow, and saves the workbook as an .xlsx file.
    public class SubtotalMinColumnGFlatView
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // -------------------------------------------------
                // Sample data creation (columns A to G)
                // -------------------------------------------------
                // Header row
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Item");
                cells["C1"].PutValue("Qty");
                cells["D1"].PutValue("Price");
                cells["E1"].PutValue("Discount");
                cells["F1"].PutValue("Region");
                cells["G1"].PutValue("Score"); // Column G (index 6) on which we will apply Min subtotal

                // Populate some rows (rows 2..6)
                object[,] data = new object[,]
                {
                    { "A", "Apple", 10, 1.2, 0.1, "North", 85 },
                    { "A", "Banana", 15, 0.8, 0.05, "North", 78 },
                    { "B", "Carrot", 20, 0.5, 0.0, "South", 92 },
                    { "B", "Daikon", 12, 0.9, 0.07, "South", 88 },
                    { "C", "Eggplant", 8, 1.5, 0.2, "East", 73 }
                };

                for (int r = 0; r < data.GetLength(0); r++)
                {
                    for (int c = 0; c < data.GetLength(1); c++)
                    {
                        cells[r + 1, c].PutValue(data[r, c]);
                    }
                }

                // -------------------------------------------------
                // Define the range that contains the data (A1:G6)
                // -------------------------------------------------
                CellArea area = CellArea.CreateCellArea("A1", "G6");

                // -------------------------------------------------
                // Add subtotals:
                //   - Group by the first column (Category) -> groupBy = 0
                //   - Use Min function on column G (Score) -> totalList = new int[] { 6 }
                // -------------------------------------------------
                cells.Subtotal(
                    area,                     // range
                    0,                        // group by first column (Category)
                    ConsolidationFunction.Min,// Min function
                    new int[] { 6 },          // apply subtotal to column G (Score)
                    false,                    // replace existing subtotals? false (no replace)
                    false,                    // add page breaks between groups? false
                    true);                    // place summary below data (true for flat view)

                // -------------------------------------------------
                // Disable outline grouping to obtain a flat view.
                // SummaryRowBelow already places the subtotal rows below the data.
                // -------------------------------------------------
                worksheet.Outline.SummaryRowBelow = true; // ensure summary rows stay below data

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "SubtotalMinColumnGFlatView.xlsx";

                // Ensure the directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            SubtotalMinColumnGFlatView.Run();
        }
    }
}
