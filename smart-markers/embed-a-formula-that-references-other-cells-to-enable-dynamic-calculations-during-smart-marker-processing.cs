// Title: Add a Formula that References Smart Marker Cells for Dynamic Calculations in Aspose.Cells (C#)
// Description: This example shows how to create a workbook, place smart markers in cells A2 and B2, embed a formula in C2 that multiplies those cells, bind a List<DataItem> as the data source, process the markers with WorkbookDesigner, recalculate formulas, and save the file as SmartMarkerWithFormula.xlsx.
// Keywords: Aspose.Cells smart marker formula | C# embed Excel formula after smart markers | WorkbookDesigner calculate formulas | dynamic calculation Aspose.Cells .NET | smart marker data source list | Excel formula referencing smart marker cells | Aspose.Cells example C# | global Aspose.Cells tutorial
// Common Searches: how to set a formula that uses smart marker values in Aspose.Cells C# | calculate formulas after WorkbookDesigner.Process() | embed Excel formula with smart markers Aspose.Cells | C# Aspose.Cells dynamic calculations with smart markers | force formula evaluation after smart marker processing
// Developer Intent: I need to embed an Excel formula that automatically uses the values filled by smart markers and have it recomputed after the markers are processed.
// Use Cases: Generate a pricing sheet where each row’s total = quantity × unit price, with quantity and price supplied by smart markers. | Create a financial summary that calculates subtotals and grand totals after expense categories and amounts are populated via smart markers. | Build an inventory report that multiplies stock count and unit cost (both from smart markers) to show total value per item.
// AI Prompts: Provide C# code that adds a formula referencing smart‑marker cells and ensures the formula is evaluated after WorkbookDesigner.Process() in Aspose.Cells. | Show how to apply the same formula to multiple rows when using a list as a smart‑marker data source. | Explain the steps to force formula recalculation and save the workbook with computed results after smart marker processing.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerFormulaDemo
{
    // Simple data class used as a data source for smart markers
    // This example shows how to create a workbook, place smart markers in cells A2 and B2, embed a formula in C2 that multiplies those cells, bind a List<DataItem> as the data source, process the markers with WorkbookDesigner, recalculate formulas, and save the file as SmartMarkerWithFormula.xlsx.
    public class DataItem
    {
        public string Name { get; set; } = string.Empty;
        public double Value { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Define smart markers that will be replaced by data source values
                //    & = indicates a smart marker; the marker name is "Data"
                cells["A2"].PutValue("&=Data.Name");
                cells["B2"].PutValue("&=Data.Value");

                // 3. Embed a formula that references the cells populated by smart markers.
                //    The formula will be evaluated after the smart markers are processed.
                cells["C2"].Formula = "=A2*B2";

                // 4. Prepare a data source (list of DataItem objects)
                List<DataItem> data = new List<DataItem>
                {
                    new DataItem { Name = "ProductA", Value = 12.5 },
                    new DataItem { Name = "ProductB", Value = 8.0 }
                };

                // 5. Create a WorkbookDesigner, assign the workbook and set the data source
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("Data", data);

                // 6. Process the smart markers (populate A2 and B2 with data)
                designer.Process();

                // 7. Calculate all formulas so that C2 reflects the computed result
                workbook.CalculateFormula();

                // 8. Save the workbook (lifecycle rule: save)
                workbook.Save("SmartMarkerWithFormula.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
