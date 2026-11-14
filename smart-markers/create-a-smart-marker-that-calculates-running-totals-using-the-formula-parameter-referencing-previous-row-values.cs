// Title: Generate an Aspose.Cells .NET smart marker that computes a running total using a Formula referencing the previous row
// Common Searches: Aspose.Cells C# smart marker formula for cumulative sum across rows | How to reference the previous row in an Aspose.Cells smart marker IF function | Create a running total column in Excel using WorkbookDesigner smart markers | Smart marker example with IF condition to add previous row value in Aspose.Cells | C# code to generate cumulative totals using smart marker Formula parameter
// Tags: Aspose.Cells smart marker total aggregation formula | WorkbookDesigner smart marker row offset reference | Excel smart marker IF expression usage | C# Aspose.Cells data source integration with smart markers | smart marker Formula parameter example

using System;
using System.Data;
using Aspose.Cells;

// The example creates a workbook, adds headers, inserts a smart marker in column A to populate values from a DataTable, and places a smart marker in column B with a Formula parameter that calculates a running total by referencing the current and prior rows. WorkbookDesigner processes the markers and the file is saved as RunningTotal.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Report";

            // Set column headers
            sheet.Cells["A1"].PutValue("Value");
            sheet.Cells["B1"].PutValue("Running Total");

            // Smart marker for the source values (column A)
            sheet.Cells["A2"].PutValue("&=Data.Value");

            // Smart marker with a formula that calculates the running total (column B)
            // IF({Row}=2, A{Row}, B{Row-1}+A{Row})
            sheet.Cells["B2"].PutValue("&=Data.Value?Formula=IF({Row}=2, A{Row}, B{Row-1}+A{Row})");

            // Prepare the data source
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("Value", typeof(double));
            dt.Rows.Add(100);
            dt.Rows.Add(150);
            dt.Rows.Add(200);
            dt.Rows.Add(250);

            // Process smart markers using WorkbookDesigner
            var designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", dt);
            designer.Process();

            // Save the workbook
            workbook.Save("RunningTotal.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
