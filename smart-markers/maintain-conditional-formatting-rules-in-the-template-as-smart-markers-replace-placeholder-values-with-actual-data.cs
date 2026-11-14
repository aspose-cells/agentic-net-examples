// Title: Keep existing conditional formatting when filling smart markers in an Excel template with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# preserve conditional formatting after WorkbookDesigner.Process | How to retain Excel conditional formatting when using smart markers in .NET | Processing smart markers without removing conditional formatting rules in Aspose.Cells | Verify conditional formatting count after filling smart markers with Aspose.Cells | Load template with smart markers and keep conditional formatting in C#
// Tags: Aspose.Cells workbookdesigner conditional formatting | smart markers data source processing C# | preserve Excel conditional rules Aspose.Cells | template workbook smart markers .NET | conditional formatting count verification

using Aspose.Cells;
using System;
using System.Data;

// The example loads an Excel template that already contains smart markers and conditional formatting, binds a DataTable as the data source, processes the smart markers with WorkbookDesigner, confirms that the original conditional formatting rules remain by outputting their count per worksheet, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load the workbook template that already contains smart markers and conditional formatting rules
        Workbook workbook = new Workbook("Template.xlsx");

        // Create a data source that matches the smart markers in the template
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Salary", typeof(double));
        dt.Rows.Add("John", 50000);
        dt.Rows.Add("Jane", 75000);
        dt.Rows.Add("Bob", 30000);

        // Attach the data source to the workbook designer
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);

        // Process smart markers – placeholder values are replaced with actual data
        designer.Process();

        // Conditional formatting rules defined in the template are preserved automatically.
        // (Optional) Verify that the rules still exist after processing.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet '{sheet.Name}' has {sheet.ConditionalFormattings.Count} conditional formatting rule(s).");
        }

        // Save the resulting workbook
        workbook.Save("Result.xlsx");
    }
}
