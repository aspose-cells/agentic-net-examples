// Title: Rename a slicer title to "Region Filter" in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to set the Title property of a slicer linked to a pivot table to 'Region Filter' and save the workbook. | Programmatically change the caption of an Excel slicer to 'Region Filter' using the Aspose.Cells Slicers API in .NET. | Update the slicer label in a generated workbook by assigning a new title and persisting the file with Aspose.Cells for C#.
// Common Searches: aspocells c# change slicer title to region filter | how to set slicer caption programmatically with Aspose.Cells .NET | rename pivot table slicer label in Excel using Aspose.Cells C# | update slicer title after creating workbook with Aspose.Cells | C# code to modify slicer title in generated Excel file
// Tags: Aspose.Cells set slicer title C# | C# modify Excel slicer caption with Aspose | Aspose.Cells rename pivot table slicer label | update slicer title programmatically Aspose.Cells | Excel slicer title change .NET

using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

// Creates a workbook, adds sample data and a pivot table, inserts a slicer for the Region field, changes the slicer title to 'Region Filter', and saves the file as SlicerTitleUpdated.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Region");
        worksheet.Cells["A2"].PutValue("North");
        worksheet.Cells["A3"].PutValue("South");
        worksheet.Cells["A4"].PutValue("East");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["B3"].PutValue(150);
        worksheet.Cells["B4"].PutValue(200);

        // Add a pivot table based on the sample data
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Region as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field
        pivotTable.CalculateData();

        // Add a slicer linked to the "Region" field of the pivot table
        int slicerIndex = worksheet.Slicers.Add(pivotTable, "F1", "Region");
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Change the slicer title to "Region Filter"
        slicer.Title = "Region Filter";

        // Save the updated workbook
        workbook.Save("SlicerTitleUpdated.xlsx");
    }
}
