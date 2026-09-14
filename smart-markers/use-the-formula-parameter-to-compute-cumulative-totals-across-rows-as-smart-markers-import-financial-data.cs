// Title: Create a cumulative sum column with Aspose.Cells smart markers and a repeated SUM formula in C#
// AI Prompts: Generate C# code that loads a DataTable into a workbook using Aspose.Cells smart markers and adds a cumulative sum column with a repeated SUM formula. | Illustrate how to enable the WorkbookDesigner.RepeatFormulasWithSubtotal property so the formula propagates to each inserted row. | Provide the sequence to recalculate all formulas after processing smart markers and export the result to an .xlsx file in C#.
// Common Searches: C# Aspose.Cells example for creating a running total column with smart markers | How to use RepeatFormulasWithSubtotal in WorkbookDesigner to repeat formulas per row | Calculate cumulative sums in Excel using smart markers and SUM formula in .NET
// Tags: smart markers cumulative total calculation | WorkbookDesigner repeat formulas flag | Aspose.Cells import DataTable C# | post‑processing formula evaluation with Aspose.Cells | C# Excel cumulative total column generation

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, defines headers and a template row with smart markers, sets a SUM formula for a cumulative total column, loads financial data from a DataTable, processes the smart markers with WorkbookDesigner while repeating the formula for each generated row, recalculates all formulas, and saves the workbook as an Excel file.
class CumulativeTotalsSmartMarkers
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells[0, 0].PutValue("Date");
            cells[0, 1].PutValue("Description");
            cells[0, 2].PutValue("Amount");
            cells[0, 3].PutValue("Cumulative Total");

            // Template row with smart markers for data import
            cells[1, 0].PutValue("&=$Date");
            cells[1, 1].PutValue("&=$Description");
            cells[1, 2].PutValue("&=$Amount");

            // Set the cumulative total formula in the first data row.
            // The formula will be repeated for each generated row.
            cells[1, 3].SetFormula("=SUM($C$2:C2)", null);

            // Prepare sample financial data in a DataTable
            DataTable dt = new DataTable("FinancialData");
            dt.Columns.Add("Date", typeof(DateTime));
            dt.Columns.Add("Description", typeof(string));
            dt.Columns.Add("Amount", typeof(double));

            dt.Rows.Add(new DateTime(2023, 1, 1), "Revenue", 1500.0);
            dt.Rows.Add(new DateTime(2023, 1, 2), "Expense", -300.0);
            dt.Rows.Add(new DateTime(2023, 1, 3), "Revenue", 1200.0);
            dt.Rows.Add(new DateTime(2023, 1, 4), "Expense", -500.0);
            dt.Rows.Add(new DateTime(2023, 1, 5), "Revenue", 800.0);

            // Process smart markers using WorkbookDesigner (lifecycle: load & process)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            // Ensure the formula in the cumulative column repeats for each generated row
            designer.RepeatFormulasWithSubtotal = true;
            designer.SetDataSource(dt);
            designer.Process();

            // Calculate all formulas after data insertion
            workbook.CalculateFormula();

            // Save the result (lifecycle: save)
            string outputPath = "CumulativeTotalsSmartMarkers.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
