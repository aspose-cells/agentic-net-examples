// Title: Delete a PivotTable from an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells remove pivot table by index from worksheet | How to delete the first PivotTable in an Excel file with Aspose.Cells | Aspose.Cells .NET code to programmatically delete a PivotTable | Remove PivotTable from worksheet using Aspose.Cells PivotTables.RemoveAt | Delete pivot tables in Excel using C# Aspose.Cells API
// Tags: Aspose.Cells PivotTables.RemoveAt example | delete Excel pivot table C# | remove pivot table by index Aspose.Cells | programmatic pivot table deletion Aspose.Cells | C# Excel pivot table removal

using Aspose.Cells;
using System;

// Load an existing workbook
Workbook workbook = new Workbook("input.xlsx");

// Access the target worksheet (by index or name)
Worksheet worksheet = workbook.Worksheets[0]; // or workbook.Worksheets["Sheet1"]

// Ensure there is at least one PivotTable before attempting removal
if (worksheet.PivotTables.Count > 0)
{
    // Remove the PivotTable at the desired index (e.g., the first one)
    worksheet.PivotTables.RemoveAt(0);
}

// Save the modified workbook
workbook.Save("output.xlsx");
