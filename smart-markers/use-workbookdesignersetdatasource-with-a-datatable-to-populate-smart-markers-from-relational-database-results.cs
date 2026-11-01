// Title: Using WorkbookDesigner.SetDataSource with a DataTable to fill Excel smart markers in Aspose.Cells for .NET
// Common Searches: how to bind a DataTable to WorkbookDesigner for smart marker processing in Aspose.Cells C# | populate Excel smart markers from a relational DataTable using Aspose.Cells .NET | Aspose.Cells example set data source for smart markers from in‑memory DataTable | C# generate Excel report with smart markers and ListObject table using WorkbookDesigner | WorkbookDesigner SetDataSource sample code for employee data in Aspose.Cells
// Tags: WorkbookDesigner SetDataSource with DataTable | Aspose.Cells smart marker data binding | C# Excel template population using smart markers | Aspose.Cells ListObject table creation | Excel report generation from relational data .NET

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates or loads an Excel template containing smart marker placeholders, builds an in‑memory DataTable with employee records, assigns the DataTable to a WorkbookDesigner via SetDataSource, processes the smart markers to expand the data into the worksheet, and saves the populated workbook as Output.xlsx.
class SmartMarkerDemo
{
    static void Main()
    {
        const string templatePath = "SmartMarkersTemplate.xlsx";
        const string outputPath = "Output.xlsx";

        try
        {
            // Ensure the template file exists; create a simple one if it does not.
            if (!File.Exists(templatePath))
            {
                CreateTemplateWorkbook(templatePath);
            }

            // Load the Excel template that contains smart markers
            Workbook workbook = new Workbook(templatePath);

            // Create a DataTable and fill it with sample relational data
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Department", typeof(string));
            dt.Columns.Add("Salary", typeof(decimal));

            dt.Rows.Add(1, "John Doe", "Finance", 75000);
            dt.Rows.Add(2, "Jane Smith", "HR", 68000);
            dt.Rows.Add(3, "Bob Johnson", "IT", 82000);
            dt.Rows.Add(4, "Alice Brown", "Marketing", 71000);

            // Initialize the WorkbookDesigner for smart marker processing
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set the DataTable as the data source for the smart markers
            designer.SetDataSource(dt);

            // Process the smart markers and populate the worksheet with data
            designer.Process();

            // Save the populated workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Creates a minimal template workbook containing smart markers for the Employees table
    private static void CreateTemplateWorkbook(string path)
    {
        try
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Employees";

            // Header row
            ws.Cells["A1"].PutValue("ID");
            ws.Cells["B1"].PutValue("Name");
            ws.Cells["C1"].PutValue("Department");
            ws.Cells["D1"].PutValue("Salary");

            // Smart marker row (starts from row 2)
            ws.Cells["A2"].PutValue("&=Employees.ID");
            ws.Cells["B2"].PutValue("&=Employees.Name");
            ws.Cells["C2"].PutValue("&=Employees.Department");
            ws.Cells["D2"].PutValue("&=Employees.Salary");

            // Add a table covering the header and smart marker rows
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, isTableStyle (true to apply default style)
            int tableIndex = ws.ListObjects.Add(0, 0, 1, 3, true);
            ListObject table = ws.ListObjects[tableIndex];
            table.TableStyleType = TableStyleType.None; // No built‑in style
            table.ShowHeaderRow = true;
            // Total row is optional; omitted here as the property may not be available in older versions

            wb.Save(path);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to create template workbook: " + ex.Message);
        }
    }
}
