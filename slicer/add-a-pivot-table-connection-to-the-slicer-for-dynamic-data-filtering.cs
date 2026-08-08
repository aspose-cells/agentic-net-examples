// Title: Add and Connect a Slicer to a Pivot Table in C# using Aspose.Cells
// Description: Demonstrates how to create a workbook, populate it with sample data, build a pivot table, add a slicer for the "Category" field, link the slicer to the pivot table for live filtering, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells slicer | C# pivot table slicer connection | dynamic Excel filtering Aspose | AddPivotConnection example | programmatic slicer pivot Aspose.Cells | Excel dashboard C# Aspose
// Common Searches: Aspose.Cells connect slicer to pivot table C# | How to add slicer for pivot field using Aspose.Cells | Dynamic filtering with slicer in Aspose.Cells .NET | C# code sample slicer pivot connection Aspose | Create interactive Excel report with slicer Aspose.Cells
// Developer Intent: Programmatically attach a slicer to a pivot table so that slicer selections instantly filter the pivot data.
// Use Cases: Build an Excel sales dashboard where users can filter categories via a slicer. | Generate reusable report templates with interactive slicer‑pivot links for end‑user analysis. | Automate workbook creation that includes pivot tables and slicers for distribution to stakeholders.
// AI Prompts: Generate C# code with Aspose.Cells that adds a slicer to a pivot table and connects it for real‑time filtering. | Show how to refresh a pivot table after adding a slicer‑pivot connection and save the workbook. | Explain positioning, sizing, and styling a slicer after linking it to a pivot table in Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Demonstrates how to create a workbook, populate it with sample data, build a pivot table, add a slicer for the "Category" field, link the slicer to the pivot table for live filtering, and save the file with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate worksheet with sample data
        cells["A1"].Value = "Category";
        cells["B1"].Value = "Product";
        cells["C1"].Value = "Sales";

        cells["A2"].Value = "Electronics";
        cells["B2"].Value = "Laptop";
        cells["C2"].Value = 1200;

        cells["A3"].Value = "Electronics";
        cells["B3"].Value = "Phone";
        cells["C3"].Value = 800;

        cells["A4"].Value = "Furniture";
        cells["B4"].Value = "Chair";
        cells["C4"].Value = 150;

        // Add a pivot table based on the data range
        PivotTableCollection pivots = sheet.PivotTables;
        int pivotIndex = pivots.Add("A1:C4", "E1", "SalesPivot");
        PivotTable pivot = pivots[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Column, "Product");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer for the "Category" field
        int slicerIndex = sheet.Slicers.Add(pivot, "G1", "Category");
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Connect the slicer to the pivot table for dynamic filtering
        slicer.AddPivotConnection(pivot);

        // Save the workbook
        workbook.Save("PivotSlicerConnection.xlsx");
    }
}
