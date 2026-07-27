using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerStyleUpdateDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];

                // Populate sample data for a pivot table
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Amount");
                dataSheet.Cells["A2"].PutValue("A");
                dataSheet.Cells["B2"].PutValue(120);
                dataSheet.Cells["A3"].PutValue("B");
                dataSheet.Cells["B3"].PutValue(80);
                dataSheet.Cells["A4"].PutValue("C");
                dataSheet.Cells["B4"].PutValue(150);
                dataSheet.Cells["A5"].PutValue("A");
                dataSheet.Cells["B5"].PutValue(30);
                dataSheet.Cells["A6"].PutValue("B");
                dataSheet.Cells["B6"].PutValue(70);
                dataSheet.Cells["A7"].PutValue("C");
                dataSheet.Cells["B7"].PutValue(20);

                // Add a pivot table based on the data
                int pivotIndex = dataSheet.PivotTables.Add("A1:B7", "D3", "PivotTable1");
                PivotTable pivot = dataSheet.PivotTables[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table for the "Category" field
                int slicerIndex = dataSheet.Slicers.Add(pivot, "F3", "Category");
                Slicer slicer = dataSheet.Slicers[slicerIndex];

                // Determine total amount from the source data (since PivotTable does not expose RowCount)
                double totalAmount = 0;
                // Data rows are 2‑7 (zero‑based rows 1‑6) in column B (index 1)
                for (int row = 1; row <= 6; row++)
                {
                    object val = dataSheet.Cells[row, 1].Value;
                    if (val is double d) totalAmount += d;
                    else if (val is int i) totalAmount += i;
                }

                // Apply slicer style based on the total amount
                if (totalAmount > 300)
                {
                    // High total – use a dark style
                    slicer.StyleType = SlicerStyleType.SlicerStyleDark2;
                }
                else if (totalAmount > 200)
                {
                    // Medium total – use a medium light style
                    slicer.StyleType = SlicerStyleType.SlicerStyleLight3;
                }
                else
                {
                    // Low total – use the default light style
                    slicer.StyleType = SlicerStyleType.SlicerStyleLight1;
                }

                // Optional: set additional slicer properties for better visibility
                slicer.Caption = "Category Filter";
                slicer.NumberOfColumns = 1;
                slicer.Shape.Width = 150;   // Use Shape.Width instead of obsolete WidthPixel
                slicer.Shape.Height = 100;  // Use Shape.Height instead of obsolete HeightPixel

                // Save the workbook (ensure the directory exists)
                string outputPath = "SlicerStyleUpdated.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}