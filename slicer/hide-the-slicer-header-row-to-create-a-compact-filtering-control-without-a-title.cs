// Title: How to hide the slicer header (caption) in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate an Excel workbook with a pivot table and add a slicer whose caption is hidden using Aspose.Cells for .NET. | Write C# code that links a slicer to a pivot table and sets the ShowCaption property to false for a compact UI. | Apply a light slicer style, configure a single column layout, and remove the header row of the slicer with Aspose.Cells.
// Common Searches: aspnet hide slicer caption in Excel workbook using Aspose.Cells | c# Aspose.Cells remove slicer header row from pivot table slicer | how to create compact slicer without title in Excel with Aspose.Cells .NET | set ShowCaption false for slicer Aspose.Cells example | customize slicer style and hide header in C# Aspose.Cells
// Tags: slicer ShowCaption property Aspose.Cells | hide slicer caption C# Aspose.Cells | compact slicer style customization .NET | pivot table slicer creation Aspose.Cells | Excel workbook slicer without caption

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    // The example creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the pivot, disables the slicer caption by setting ShowCaption to false, applies a light style, sets a single column layout, and saves the file as SlicerWithoutHeader.xlsx.
    public class HideSlicerHeader
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for a pivot table
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Amount";
                cells["A2"].Value = "Fruit";
                cells["B2"].Value = 120;
                cells["A3"].Value = "Vegetable";
                cells["B3"].Value = 80;
                cells["A4"].Value = "Grain";
                cells["B4"].Value = 50;

                // Add a pivot table based on the data range
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh pivot cache and calculate data
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table
                int slicerIdx = sheet.Slicers.Add(pivot, "F1", "Category");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // Hide the slicer header (caption) to make it compact
                slicer.ShowCaption = false;

                // Optional: adjust appearance
                slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
                slicer.NumberOfColumns = 1;

                // Save the workbook
                string outputPath = "SlicerWithoutHeader.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            HideSlicerHeader.Run();
        }
    }
}
