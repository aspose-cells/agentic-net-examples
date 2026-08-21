// Title: C# – Dynamically Change Aspose.Cells Slicer Style Based on Pivot Total
// Description: Creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the Category field, calculates the total amount, and sets the slicer's StyleType to a dark or light preset depending on whether the total exceeds a threshold. The slicer is refreshed and the file saved.
// Keywords: Aspose.Cells | C# | slicer style | conditional slicer formatting | pivot table slicer | SlicerStyleType | dynamic Excel styling | programmatic slicer update | Excel automation | calculate formula Aspose
// Common Searches: change slicer style Aspose.Cells C# | conditional slicer formatting based on cell value | set slicer style programmatically .NET | apply dark slicer style when total > 50 Aspose | refresh slicer after style change Aspose.Cells
// Developer Intent: Apply a conditional style to an Aspose.Cells slicer by reading a calculated cell value and setting the Slicer.StyleType property accordingly.
// Use Cases: Highlight a slicer with a dark theme when the sum of a column exceeds a business threshold. | Revert to a light slicer style for lower totals to improve visual contrast. | Automatically refresh the slicer after changing its style to reflect the new appearance in the workbook.
// AI Prompts: Generate C# code that updates an Aspose.Cells slicer style based on a pivot table total. | Show how to conditionally apply dark or light slicer styles in Aspose.Cells using a calculated cell value. | Explain the steps to refresh a slicer after modifying its StyleType in Aspose.Cells .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerStyleUpdateDemo
{
    // Creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the Category field, calculates the total amount, and sets the slicer's StyleType to a dark or light preset depending on whether the total exceeds a threshold. The slicer is refreshed and the file saved.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for the pivot table
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Amount");
                cells["A2"].PutValue("A");
                cells["B2"].PutValue(5);
                cells["A3"].PutValue("B");
                cells["B3"].PutValue(15);
                cells["A4"].PutValue("C");
                cells["B4"].PutValue(25);
                cells["A5"].PutValue("D");
                cells["B5"].PutValue(35);

                // Add a cell that will drive the style decision (e.g., total amount)
                cells["C1"].PutValue("Total");
                cells["C2"].Formula = "SUM(B2:B5)";

                // Ensure formulas are calculated before reading their values
                workbook.CalculateFormula();

                // Create a pivot table based on the data
                int pivotIdx = sheet.PivotTables.Add("A1:B5", "E1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table's Category field
                int slicerIdx = sheet.Slicers.Add(pivot, "G1", "Category");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // Determine the style based on the driven criteria (total amount > 50 => dark style)
                double totalAmount = cells["C2"].DoubleValue;
                if (totalAmount > 50)
                {
                    slicer.StyleType = SlicerStyleType.SlicerStyleDark2; // Dark style for higher totals
                }
                else
                {
                    slicer.StyleType = SlicerStyleType.SlicerStyleLight1; // Light style otherwise
                }

                // Refresh the slicer to apply any changes
                slicer.Refresh();

                // Save the workbook
                workbook.Save("SlicerStyleUpdated.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
