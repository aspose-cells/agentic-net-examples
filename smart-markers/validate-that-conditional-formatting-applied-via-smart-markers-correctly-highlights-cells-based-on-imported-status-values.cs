// Title: Validate that conditional formatting applied via smart markers correctly highlights employee status cells using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel template with smart markers, binds a DataTable of employee names and status values, runs WorkbookDesigner.Process, then iterates the Status column to assert that each cell's foreground color matches the predefined highlight color for Approved, Pending, and Rejected, logging any mismatches. | Write a C# method that receives a processed Worksheet, reads the text in the Status column, maps each status to its expected highlight color, retrieves the actual cell style foreground color, compares the two, and returns a collection of validation results indicating success or failure for each row.
// Common Searches: how to test conditional formatting colors after applying smart markers with Aspose.Cells in C# | Aspose.Cells verify cell foreground color based on smart marker data source | C# code to compare expected and actual cell colors after WorkbookDesigner processing | checking conditional formatting results for a status column using Aspose.Cells .NET
// Tags: smart markers conditional formatting verification | WorkbookDesigner status cell color validation | Aspose.Cells compare cell foreground color | C# data-driven conditional formatting test | Excel conditional formatting validation Aspose.Cells

using System;
using System.Data;
using System.Drawing;
using Aspose.Cells;

// The example loads a workbook containing smart markers and conditional formatting rules, populates a DataTable with employee names and status values, processes the smart markers via WorkbookDesigner, then iterates through the Status column to compare each cell's actual foreground color with the expected highlight color (Approved = LightGreen, Pending = Yellow, Rejected = LightSalmon), reporting any discrepancies before saving the result workbook.
class Program
{
    static void Main()
    {
        // Load the template workbook that contains smart markers and conditional formatting rules
        Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

        // Prepare a data source with status values
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Status", typeof(string));
        dt.Rows.Add("John", "Approved");
        dt.Rows.Add("Jane", "Pending");
        dt.Rows.Add("Bob", "Rejected");

        // Apply smart markers using WorkbookDesigner
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);
        designer.Process();

        // Validate that conditional formatting highlights cells according to the Status values
        Worksheet ws = workbook.Worksheets[0]; // assuming data is on the first sheet

        // Assuming the Status column is column B (index 1) and data starts at row 2 (index 1)
        for (int row = 1; row <= dt.Rows.Count; row++)
        {
            string status = ws.Cells[row, 1].StringValue; // read the status text
            Style cellStyle = ws.Cells[row, 1].GetStyle();
            Color actualColor = cellStyle.ForegroundColor; // color applied by conditional formatting

            // Determine the expected highlight color for each status
            Color expectedColor = Color.Empty;
            switch (status)
            {
                case "Approved":
                    expectedColor = Color.LightGreen;
                    break;
                case "Pending":
                    expectedColor = Color.Yellow;
                    break;
                case "Rejected":
                    expectedColor = Color.LightSalmon;
                    break;
            }

            // Compare actual and expected colors
            if (actualColor.ToArgb() != expectedColor.ToArgb())
            {
                Console.WriteLine($"Row {row + 1}: Status '{status}' has incorrect highlight. Expected {expectedColor}, got {actualColor}.");
            }
            else
            {
                Console.WriteLine($"Row {row + 1}: Status '{status}' correctly highlighted.");
            }
        }

        // Save the processed workbook
        workbook.Save("Result.xlsx");
    }
}
