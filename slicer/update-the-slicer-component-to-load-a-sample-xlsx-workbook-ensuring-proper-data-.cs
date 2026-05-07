using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace SlicerRefreshExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains a PivotTable and a Slicer
            // (uses the provided Workbook(string) constructor – lifecycle rule)
            string inputPath = "SampleData.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Assume the slicer is placed on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Access the slicer collection (provided property)
            SlicerCollection slicers = sheet.Slicers;

            // Ensure there is at least one slicer; otherwise add one for demonstration
            if (slicers.Count == 0)
            {
                // Create a simple PivotTable to bind the slicer to
                // (uses the provided PivotTables.Add method)
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "DemoPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, 0);
                pivot.AddFieldToArea(PivotFieldType.Data, 1);
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the first field of the pivot table
                // (uses the SlicerCollection.Add method)
                int slicerIndex = slicers.Add(pivot, "F1", pivot.BaseFields[0]);
                Slicer slicer = slicers[slicerIndex];
                // Refresh the newly added slicer
                slicer.Refresh();
            }
            else
            {
                // Refresh each existing slicer to ensure data binding is up‑to‑date
                foreach (Slicer slicer in slicers)
                {
                    slicer.Refresh(); // uses Slicer.Refresh method
                }
            }

            // Save the workbook after refreshing slicers
            // (uses the provided Workbook.Save(string) method – lifecycle rule)
            string outputPath = "SampleData_Refreshed.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to '{outputPath}'. Slicers refreshed successfully.");
        }
    }
}