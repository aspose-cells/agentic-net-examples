using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the template workbook (load rule)
        Workbook workbook = new Workbook("Template.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set up header
        sheet.Cells["A1"].PutValue("Score");

        // Insert a conditional smart marker.
        // The marker uses the IF function: if Score > 80 then "Pass" else "Fail"
        sheet.Cells["A2"].PutValue("&=IF(Score>80,\"Pass\",\"Fail\")");

        // Define the range that contains smart markers and give it the required name
        Aspose.Cells.Range smartRange = sheet.Cells.CreateRange("A2");
        smartRange.Name = "_CellsSmartMarkers";

        // Prepare a data source (DataTable) with a column that matches the smart marker variable
        DataTable dt = new DataTable("Students");
        dt.Columns.Add("Score", typeof(int));
        dt.Rows.Add(92);   // Should result in "Pass"
        dt.Rows.Add(67);   // Should result in "Fail"
        dt.Rows.Add(81);   // Should result in "Pass"

        // Create a WorkbookDesigner, assign the workbook and set the data source
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook
        };
        designer.SetDataSource(dt);

        // Process the smart markers (process rule)
        designer.Process(true);

        // Save the populated workbook (save rule)
        workbook.Save("Result.xlsx");
    }
}