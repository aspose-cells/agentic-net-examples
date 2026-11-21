// Title: Generate cumulative running totals in Excel using the smart markers Formula parameter with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# smart markers cumulative sum using Formula placeholder | How to create a running total column with WorkbookDesigner and smart markers | C# example of importing a financial DataTable into Excel and calculating cumulative amount | Using {Row} placeholder in smart marker formula for cumulative totals in Aspose.Cells | Smart markers formula for summing previous rows in Excel via Aspose.Cells .NET
// Tags: smart markers cumulative sum formula | WorkbookDesigner import DataTable | Aspose.Cells running total column | Excel template smart markers {Row} placeholder | C# generate cumulative totals Excel

using System;
using System.Data;
using Aspose.Cells;

// This example loads an Excel template that contains smart markers, builds a DataTable with Date and Amount fields, assigns it to a WorkbookDesigner, and uses the Formula parameter "=SUM($B$2:B{Row})" in a smart marker to automatically compute cumulative totals in the third column for each imported row. The processed workbook is then saved as Result.xlsx.
class Program
{
    static void Main()
    {
        // Load the Excel template that contains smart markers.
        // The template should have the following smart markers in row 2 (header row in row 1):
        //   A2:  <=Data.Date>
        //   B2:  <=Data.Amount>
        //   C2:  <=Data.Cumulative,Formula="=SUM($B$2:B{Row})">
        // The Formula parameter uses the {Row} placeholder which Aspose.Cells replaces with the current row number,
        // thus each row in column C will show the cumulative sum of the Amount column up to that row.
        Workbook workbook = new Workbook("Template.xlsx");

        // Prepare the financial data source.
        DataTable dt = new DataTable("Data");
        dt.Columns.Add("Date", typeof(DateTime));
        dt.Columns.Add("Amount", typeof(double));

        // Sample data rows.
        dt.Rows.Add(new DateTime(2023, 1, 1), 1000);
        dt.Rows.Add(new DateTime(2023, 2, 1), 1500);
        dt.Rows.Add(new DateTime(2023, 3, 1), 2000);
        dt.Rows.Add(new DateTime(2023, 4, 1), 1200);

        // Set up the WorkbookDesigner and assign the data source.
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);

        // Process the smart markers. The Formula parameter will automatically generate
        // cumulative totals in column C for each imported row.
        designer.Process();

        // Save the populated workbook.
        workbook.Save("Result.xlsx");
    }
}
