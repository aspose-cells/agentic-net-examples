using System;
using Aspose.Cells;

class MergeJsonDataSourceDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a smart marker that will consume the merged JSON collection
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("&=$Employees.Name");

        // Two separate JSON arrays that we want to merge
        string jsonArray1 = "[{\"Name\":\"John\"},{\"Name\":\"Jane\"}]";
        string jsonArray2 = "[{\"Name\":\"Bob\"},{\"Name\":\"Alice\"}]";

        // Merge the arrays into a single JSON array string
        // Remove the outer brackets from each array and concatenate with a comma
        string mergedJson = "[" +
            jsonArray1.TrimStart('[').TrimEnd(']') + "," +
            jsonArray2.TrimStart('[').TrimEnd(']') +
            "]";

        // Set up the WorkbookDesigner and assign the workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;

        // Register the merged JSON as a data source named "Employees"
        designer.SetJsonDataSource("Employees", mergedJson);

        // Process the smart markers to populate the data
        designer.Process();

        // Save the resulting workbook
        workbook.Save("MergedJsonSmartMarker.xlsx");
    }
}