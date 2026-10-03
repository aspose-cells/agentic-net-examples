// Title: Generate a running cumulative total column with the Formula parameter in Aspose.Cells smart markers (C#)
// AI Prompts: Create an Excel workbook in C# that uses smart markers to insert Date, Description, Amount and a Formula smart marker that calculates a running aggregate for the Amount column. | Modify the smart marker formula to compute aggregate totals for a different column or start row, and update the WorkbookDesigner data source accordingly. | Add category grouping to the data and generate separate subtotal columns per category using smart markers and the Formula parameter.
// Common Searches: how to calculate running total with Aspose.Cells smart markers C# | Aspose.Cells Formula parameter cumulative sum example | C# smart markers SUM(C$2:C{Row}) for financial data | using WorkbookDesigner to create cumulative total column in Excel | smart markers cumulative total per row Aspose.Cells tutorial
// Tags: Formula syntax SUM range implementation | data source binding Aspose.Cells | C# Excel export cumulative calculation | financial data import Aspose.Cells example | Aspose.Cells progressive sum via formula parameter

using System;
using System.Data;
using Aspose.Cells;

// The example creates a new workbook, adds headers, places smart markers for Date, Description, Amount, and a Formula marker that computes a cumulative sum of the Amount column, binds a DataTable as the data source, processes the markers with WorkbookDesigner, and saves the result as CumulativeTotals.xlsx.
class CumulativeTotalsSmartMarker
{
    static void Main()
    {
        try
        {
            // Create a new workbook (template) and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "FinancialData";

            // Set up headers in the first row
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Description");
            sheet.Cells["C1"].PutValue("Amount");
            sheet.Cells["D1"].PutValue("Cumulative Total");

            // Insert smart markers in the second row.
            // & =Data.Date &      -> inserts the Date value from the data source.
            // & =Data.Description & -> inserts the Description.
            // & =Data.Amount &    -> inserts the Amount.
            // &Formula=SUM(C$2:C{Row})& -> computes cumulative sum of Amount column up to the current row.
            sheet.Cells["A2"].PutValue("&=Data.Date&");
            sheet.Cells["B2"].PutValue("&=Data.Description&");
            sheet.Cells["C2"].PutValue("&=Data.Amount&");
            sheet.Cells["D2"].PutValue("&Formula=SUM(C$2:C{Row})&");

            // Prepare the data source (DataTable) with financial records
            DataTable dt = new DataTable("Data");
            dt.Columns.Add("Date", typeof(DateTime));
            dt.Columns.Add("Description", typeof(string));
            dt.Columns.Add("Amount", typeof(double));

            // Sample data rows
            dt.Rows.Add(new DateTime(2023, 1, 1), "Opening Balance", 1000.00);
            dt.Rows.Add(new DateTime(2023, 1, 5), "Revenue", 2500.00);
            dt.Rows.Add(new DateTime(2023, 1, 10), "Expense", -800.00);
            dt.Rows.Add(new DateTime(2023, 1, 15), "Revenue", 1200.00);
            dt.Rows.Add(new DateTime(2023, 1, 20), "Expense", -500.00);

            // Use WorkbookDesigner to process smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process(); // Populate smart markers with data

            // Save the resulting workbook
            workbook.Save("CumulativeTotals.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
