// Title: Using an IF smart marker to display Pass or Fail based on a numeric Score with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# conditional smart marker IF function example | generate Pass/Fail column using smart markers in Excel with Aspose.Cells | set up WorkbookDesigner with DataTable and conditional smart marker in C# | use IF expression in Aspose.Cells smart marker to evaluate score values | create conditional smart markers for pass fail logic in .NET
// Tags: IF smart marker expression in Aspose.Cells | WorkbookDesigner data source binding for smart markers | C# generate Pass Fail column using smart markers | Aspose.Cells conditional smart marker Excel output | smart marker evaluation with DataTable in .NET

using System;
using System.Data;
using Aspose.Cells;

// The example demonstrates how to create a new workbook, add header cells, insert a simple smart marker for a numeric Score, and apply a conditional smart marker using the IF function to output "Pass" when Score > 60 and "Fail" otherwise. It binds a DataTable containing Score values as the data source, processes the markers with WorkbookDesigner, and saves the resulting Excel file.
class ConditionalSmartMarkerExample
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Header row
        sheet.Cells["A1"].PutValue("Score");
        sheet.Cells["B1"].PutValue("Result");

        // Smart markers for data rows (starting at row 2)
        // Simple smart marker for the numeric score
        sheet.Cells["A2"].PutValue("&=Score");
        // Conditional smart marker: displays "Pass" if Score > 60, otherwise "Fail"
        sheet.Cells["B2"].PutValue("&=IF(Score>60,\"Pass\",\"Fail\")");

        // Prepare the data source (a DataTable named "Data")
        DataTable dt = new DataTable("Data");
        dt.Columns.Add("Score", typeof(int));
        dt.Rows.Add(85); // Example score that will be "Pass"
        dt.Rows.Add(58); // Example score that will be "Fail"
        dt.Rows.Add(73); // Example score that will be "Pass"

        // Process the smart markers (designer uses the data source)
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);
        designer.Process();

        // Save the workbook (save rule)
        workbook.Save("ConditionalSmartMarkers_Output.xlsx");
    }
}
