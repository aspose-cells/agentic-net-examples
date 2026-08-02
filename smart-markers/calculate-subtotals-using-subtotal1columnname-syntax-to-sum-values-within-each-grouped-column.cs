// Title: C# – Calculate Department‑wise Salary Subtotals in Excel with Aspose.Cells Subtotal (subtotal1:ColumnName) syntax
// Description: This example creates a workbook, adds a header and employee data, defines a cell range, and uses Cells.Subtotal to group rows by the Department column and sum the Salary column. The call replaces existing subtotals, inserts page breaks, and places a summary row below each group, then saves the file as SubtotalDemo.xlsx.
// Keywords: Aspose.Cells | C# | Excel subtotal | group by column | sum function | department salary total | smart markers | subtotal1 syntax | cells.Subtotal method | .NET Excel automation
// Common Searches: Aspose.Cells subtotal example C# | group rows and sum in Excel using Aspose.Cells | how to add department subtotals with Aspose.Cells | C# cells.Subtotal method usage | subtotal1:ColumnName smart marker syntax
// Developer Intent: Generate an Excel file that automatically groups employee rows by Department and inserts subtotal rows that sum Salary, using the Aspose.Cells Subtotal API.
// Use Cases: Produce payroll reports that show total salary per department. | Create printable Excel sheets with page breaks after each department group. | Refresh subtotals programmatically when underlying data changes.
// AI Prompts: Show how to modify the Subtotal call to also calculate average salary per department. | Provide code to export the workbook with subtotals to PDF after saving. | Explain how to achieve the same grouped sum using the (subtotal1:ColumnName) smart marker syntax in an Aspose.Cells template.

using Aspose.Cells;
using System;

// This example creates a workbook, adds a header and employee data, defines a cell range, and uses Cells.Subtotal to group rows by the Department column and sum the Salary column. The call replaces existing subtotals, inserts page breaks, and places a summary row below each group, then saves the file as SubtotalDemo.xlsx.
class SubtotalExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add header row
        cells["A1"].PutValue("Department");
        cells["B1"].PutValue("Employee");
        cells["C1"].PutValue("Salary");

        // Populate sample data
        object[,] data = new object[,]
        {
            { "HR",      "Alice",   5000 },
            { "HR",      "Bob",     4500 },
            { "IT",      "Charlie", 7000 },
            { "IT",      "David",   6500 },
            { "IT",      "Eve",     7200 },
            { "Finance", "Frank",   6000 }
        };

        for (int i = 0; i < data.GetLength(0); i++)
        {
            cells[i + 1, 0].PutValue(data[i, 0]); // Department
            cells[i + 1, 1].PutValue(data[i, 1]); // Employee
            cells[i + 1, 2].PutValue(data[i, 2]); // Salary
        }

        // Define the cell area that includes the header and data
        CellArea area = CellArea.CreateCellArea("A1", "C7");

        // Apply subtotal:
        // - Group by column 0 (Department)
        // - Use SUM function
        // - Subtotal the Salary column (index 2)
        // - Replace existing subtotals, add page breaks, place summary below data
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 }, true, true, true);

        // Save the workbook
        workbook.Save("SubtotalDemo.xlsx");
    }
}
