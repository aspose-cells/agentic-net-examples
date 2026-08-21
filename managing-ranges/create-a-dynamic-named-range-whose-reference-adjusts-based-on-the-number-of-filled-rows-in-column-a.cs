// Title: C# Aspose.Cells: Create a Dynamic Named Range that Expands with Column A Data
// Description: Shows how to use Aspose.Cells for .NET to add a workbook‑level named range defined with an OFFSET‑COUNTA formula, automatically adjusting to the count of filled rows in column A. The sample populates data, sets the RefersTo property, recalculates formulas, refreshes dynamic arrays, and saves the workbook.
// Keywords: Aspose.Cells | C# dynamic named range | OFFSET formula | COUNTA | .NET Excel automation | auto‑expanding range | named range column A | refresh dynamic array formulas | Excel workbook programmatic
// Common Searches: Aspose.Cells create dynamic named range C# | OFFSET COUNTA named range .NET | auto expanding range column A Aspose | programmatically set RefersTo property | refresh dynamic array formulas Aspose.Cells | C# Excel named range that grows with data
// Developer Intent: Programmatically define a workbook named range whose size updates automatically based on the number of non‑empty cells in column A.
// Use Cases: Drive chart data series that grow as new items are added to column A. | Populate data‑validation lists that stay current without manual range changes. | Feed formulas such as SUM or VLOOKUP with a range that always reflects the full column A dataset. | Create a pivot‑table source that expands with incoming rows. | Generate reports where the source range must adapt to varying data lengths.
// AI Prompts: Write C# Aspose.Cells code to create a dynamic named range that starts at B2 and expands across two columns. | Explain how to modify the OFFSET formula to ignore blank cells in the middle of column A. | Show how to update an existing named range after rows are inserted or deleted using Aspose.Cells. | Provide steps to refresh all dynamic array formulas after changing a named range in a workbook. | Generate a PowerShell script that uses Aspose.Cells to add a dynamic named range to an existing Excel file.

using System;
using Aspose.Cells;

// Shows how to use Aspose.Cells for .NET to add a workbook‑level named range defined with an OFFSET‑COUNTA formula, automatically adjusting to the count of filled rows in column A. The sample populates data, sets the RefersTo property, recalculates formulas, refreshes dynamic arrays, and saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Sample data in column A (optional, demonstrates the dynamic behavior)
        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue($"Item{i + 1}");
        }

        // Add a named range to the workbook
        int nameIdx = wb.Worksheets.Names.Add("DynamicRange");
        Name dynamicName = wb.Worksheets.Names[nameIdx];

        // Define the dynamic range using OFFSET and COUNTA so it expands with filled rows in column A
        // Formula: =OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)
        string dynamicFormula = $"=OFFSET({ws.Name}!$A$1,0,0,COUNTA({ws.Name}!$A:$A),1)";
        dynamicName.RefersTo = dynamicFormula;

        // Recalculate formulas and refresh any dynamic array formulas (good practice after changes)
        wb.CalculateFormula();
        wb.RefreshDynamicArrayFormulas(true);

        // Save the workbook
        wb.Save("DynamicNamedRange.xlsx");
    }
}
