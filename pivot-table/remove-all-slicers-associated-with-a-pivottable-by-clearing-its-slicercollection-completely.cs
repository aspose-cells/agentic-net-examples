// Title: Aspose.Cells for .NET – Remove All PivotTable Slicers from a Worksheet
// Description: Load an existing workbook, access the target worksheet, clear its SlicerCollection to delete every slicer linked to any PivotTable, and save the updated file using Aspose.Cells C# API.
// Keywords: Aspose.Cells clear slicers | remove pivot table slicers .NET | worksheet slicers.Clear() | delete Excel slicers programmatically | C# Aspose.Cells slicer collection | pivot table slicer removal
// Common Searches: how to delete all slicers in an Excel file with Aspose.Cells | clear slicer collection on a worksheet C# | remove pivot table slicers using Aspose.Cells for .NET | Aspose.Cells delete slicers programmatically | C# code to clear Excel slicers
// Developer Intent: Programmatically delete every slicer associated with PivotTables by clearing the worksheet’s SlicerCollection.
// Use Cases: Prepare a distribution‑ready workbook by stripping interactive slicers before publishing. | Reset a PivotTable template so end users can create fresh slicers. | Automate cleanup of temporary slicers after generating a pivot‑based analysis.
// AI Prompts: Generate C# code that lists all slicer names on a worksheet before clearing them with Aspose.Cells. | Explain how to verify that slicers have been removed after calling worksheet.Slicers.Clear(). | Show how to clear slicers for a specific PivotTable without affecting slicers linked to other PivotTables.

using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Load an existing workbook, access the target worksheet, clear its SlicerCollection to delete every slicer linked to any PivotTable, and save the updated file using Aspose.Cells C# API.
class Program
{
    static void Main()
    {
        // Load an existing workbook that contains the PivotTable and its slicers
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet that holds the PivotTable
        Worksheet worksheet = workbook.Worksheets[0];

        // Clear all slicers on this worksheet (removes slicers linked to the PivotTable)
        worksheet.Slicers.Clear();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
