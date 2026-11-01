// Title: How to add a master smart marker in an Excel template to repeat rows for each master record using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells master smart marker repeat rows for each record in template | Using WorkbookDesigner to process master smart markers and duplicate rows in Excel with Aspose.Cells | SetDataSource with DataTable for master smart markers in Aspose.Cells .NET | Insert &=MasterRecords smart marker in Excel template using Aspose.Cells C# | Generate multiple rows from a DataTable using smart markers in Aspose.Cells
// Tags: master smart marker row duplication Aspose.Cells | WorkbookDesigner SetDataSource DataTable | smart marker syntax &=MasterRecords .NET | repeat rows in Excel template using Aspose.Cells | process smart markers C#

using System;
using System.Data;
using Aspose.Cells;

// The example loads an Excel template, places master smart markers &=MasterRecords.Name and &=MasterRecords.Amount in cells A2 and B2, creates a DataTable named 'MasterRecords' with sample data, assigns it to a WorkbookDesigner via SetDataSource, processes the smart markers to automatically repeat the row for each record, and saves the populated workbook as Output.xlsx.
class SmartMarkerExample
{
    static void Main()
    {
        // Load the template workbook (replace with your template path)
        var workbook = new Workbook("Template.xlsx");

        // Get the first worksheet
        var sheet = workbook.Worksheets[0];
        var cells = sheet.Cells;

        // Place a master smart marker in the first row to repeat for each master record
        // The smart marker syntax: &=MasterRecords.FieldName
        // Here we assume the master record has fields "Name" and "Amount"
        cells["A2"].PutValue("&=MasterRecords.Name");
        cells["B2"].PutValue("&=MasterRecords.Amount");

        // Prepare sample data for the master records
        var dt = new DataTable("MasterRecords");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Amount", typeof(double));

        dt.Rows.Add("Alice", 1200.50);
        dt.Rows.Add("Bob", 950.75);
        dt.Rows.Add("Charlie", 1430.00);

        // Set the data source for the smart markers
        var designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);
        designer.Process(); // Process the smart markers and repeat rows

        // Save the resulting workbook (replace with your desired output path)
        workbook.Save("Output.xlsx");
    }
}
