using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class RenderSlicerExample
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        string inputPath = "InputData.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Assume the first worksheet contains source data for the pivot table
        Worksheet dataSheet = workbook.Worksheets[0];
        Cells cells = dataSheet.Cells;

        // If the workbook does not already have a pivot table, create one
        // Here we create a pivot table on a new worksheet for demonstration
        Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");
        // Define the data range (adjust as needed)
        string dataRange = dataSheet.Name + "!A1:C9";
        // Add the pivot table at cell C3 on the pivot sheet
        int pivotIndex = pivotSheet.PivotTables.Add(dataRange, "C3", "SalesPivot");
        PivotTable pivot = pivotSheet.PivotTables[pivotIndex];

        // Configure the pivot fields (adjust field names/indexes to match your data)
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");      // Row field
        pivot.AddFieldToArea(PivotFieldType.Column, "Year");   // Column field
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");   // Data field
        pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium10;
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer linked to the "Fruit" field of the pivot table
        // The slicer will be placed with its upper‑left corner at cell E12
        int slicerIndex = pivotSheet.Slicers.Add(pivot, "E12", "Fruit");
        Slicer slicer = pivotSheet.Slicers[slicerIndex];

        // Optional: customize slicer appearance
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
        slicer.Caption = "Fruit Filter";
        slicer.ShowCaption = true;
        slicer.NumberOfColumns = 1;

        // Refresh the slicer to ensure it reflects the current pivot data
        slicer.Refresh();

        // Save the modified workbook (replace with your desired output path)
        string outputPath = "OutputWithSlicer.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook saved with slicer at '{outputPath}'.");
    }
}