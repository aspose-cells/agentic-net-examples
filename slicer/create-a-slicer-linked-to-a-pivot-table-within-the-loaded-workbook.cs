using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class SlicerDemo
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Obtain or create a pivot table
            PivotTable pivot;
            if (worksheet.PivotTables.Count > 0)
            {
                pivot = worksheet.PivotTables[0];
            }
            else
            {
                // Create a simple pivot table from a sample range (A1:C10)
                int pivotIndex = worksheet.PivotTables.Add("=Sheet1!A1:C10", "E5", "DemoPivot");
                pivot = worksheet.PivotTables[pivotIndex];

                // Note: Aspose.Cells automatically adds default fields when a pivot table is created.
                // Additional field configuration can be added here if needed.
            }

            // Determine a field name for the slicer (use first row field if available)
            string slicerField = "Fruit";
            if (pivot.RowFields.Count > 0)
                slicerField = pivot.RowFields[0].Name;

            // Add slicer linked to the pivot table
            int slicerIndex = worksheet.Slicers.Add(pivot, "E2", slicerField);
            Slicer slicer = worksheet.Slicers[slicerIndex];
            slicer.Caption = $"{slicerField} Slicer";
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}