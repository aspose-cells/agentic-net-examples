// Title: C# – Remove All Slicers from a Worksheet and Save as XLSX with Aspose.Cells
// Description: Shows how to build a workbook, add a pivot table and a slicer, clear all slicers from the worksheet using Worksheet.Slicers.Clear (or RemoveAll), and save the result as an XLSX file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# remove slicers | clear worksheet slicers .NET | Worksheet.Slicers.Clear | delete all slicers Aspose.Cells | save workbook as XLSX | pivot table slicer removal | Aspose.Cells example
// Common Searches: Aspose.Cells delete all slicers | how to clear slicers in .NET workbook | remove slicer collection before saving Excel | C# Aspose.Cells clear slicers | Worksheet.Slicers.Clear usage
// Developer Intent: Programmatically delete every slicer from a worksheet and then export the workbook as an XLSX file using Aspose.Cells for .NET.
// Use Cases: Create a report with interactive slicers, then generate a clean version for distribution without slicer controls. | Automate a data‑pipeline that resets slicer state before archiving or sending the workbook to downstream systems. | Prepare a pivot‑driven workbook for printing or PDF conversion by removing all slicer UI elements.
// AI Prompts: Give C# code that uses Aspose.Cells to clear all slicers from a worksheet and save the file as XLSX. | Explain how to call Worksheet.Slicers.Clear (or RemoveAll) to delete every slicer before exporting an Excel workbook with Aspose.Cells. | Show the steps to programmatically remove slicers linked to a pivot table and then save the workbook in XLSX format using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerClearDemo
{
    // Shows how to build a workbook, add a pivot table and a slicer, clear all slicers from the worksheet using Worksheet.Slicers.Clear (or RemoveAll), and save the result as an XLSX file with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate worksheet with sample data for a pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Year");
                sheet.Cells["C1"].PutValue("Amount");

                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue(2020);
                sheet.Cells["C2"].PutValue(50);

                sheet.Cells["A3"].PutValue("Fruit");
                sheet.Cells["B3"].PutValue(2021);
                sheet.Cells["C3"].PutValue(70);

                sheet.Cells["A4"].PutValue("Vegetable");
                sheet.Cells["B4"].PutValue(2020);
                sheet.Cells["C4"].PutValue(30);

                sheet.Cells["A5"].PutValue("Vegetable");
                sheet.Cells["B5"].PutValue(2021);
                sheet.Cells["C5"].PutValue(60);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("=Sheet1!A1:C5", "E3", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Column, "Year");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table.
                // The destination cell must be a valid cell reference (e.g., "F2").
                SlicerCollection slicers = sheet.Slicers;
                int slicerIndex = slicers.Add(pivot, "F2", "Category");
                Slicer slicer = slicers[slicerIndex];
                slicer.Caption = "Category Filter";

                // Clear all slicers from the worksheet
                sheet.Slicers.Clear();

                // Save the workbook
                workbook.Save("WorkbookWithoutSlicers.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
