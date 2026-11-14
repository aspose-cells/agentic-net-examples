// Title: Preserving a leading apostrophe in WorkbookDesigner smart‑marker mail merge with Aspose.Cells for .NET
// Common Searches: Aspose.Cells WorkbookDesigner keep leading apostrophe in merged cell C# | set DesignerOptions PreserveLeadingApostrophe for smart markers Aspose.Cells | C# mail merge Excel output retain leading single quote using Aspose.Cells | how to prevent Aspose.Cells from stripping leading apostrophe during merge | smart marker literal apostrophe issue Aspose.Cells .NET
// Tags: WorkbookDesigner PreserveLeadingApostrophe setting | Aspose.Cells DesignerOptions leading apostrophe handling | smart marker mail merge literal apostrophe | C# Excel export preserve leading single quote | Aspose.Cells smart markers apostrophe preservation

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.Data;
using System.IO;

// The sample loads or creates an Excel workbook, inserts a smart‑marker "&=EmployeeName", supplies a DataTable where the EmployeeName value starts with an apostrophe, processes the mail merge with WorkbookDesigner, and saves the result, showing how to retain the leading apostrophe in the merged cell.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Add a simple placeholder sheet with a merge field for demonstration.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("&=EmployeeName");
            }

            // Initialize WorkbookDesigner with the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare data source containing a value with a leading apostrophe.
            DataTable dt = new DataTable();
            dt.Columns.Add("EmployeeName", typeof(string));
            dt.Rows.Add("'Alice"); // The apostrophe should appear in the final cell.

            // Assign the data source to the designer.
            designer.SetDataSource(dt);

            // Perform the mail‑merge processing.
            designer.Process();

            // Save the resulting workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
