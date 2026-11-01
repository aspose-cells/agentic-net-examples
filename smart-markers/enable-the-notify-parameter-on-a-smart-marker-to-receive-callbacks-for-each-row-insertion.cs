// Title: How to enable the Notify property on Aspose.Cells WorkbookDesigner to receive row‑insertion callbacks for smart markers in C#
// AI Prompts: Generate C# code that sets WorkbookDesigner.Notify = true and attaches a SmartMarkerEventHandler to capture each row inserted by a smart marker. | Explain how to implement a callback method that receives the inserted row index and source data when a smart marker expands rows in Aspose.Cells. | Show the complete workflow: enable notifications, register the event handler, process smart markers, and save the workbook.
// Common Searches: Aspose.Cells C# enable Notify on WorkbookDesigner for smart marker row events | how to get callback for each row added by smart markers using Aspose.Cells | C# example of WorkbookDesigner.Notify property with SmartMarkerEventHandler | receive row insertion notifications when processing smart markers in Aspose.Cells | Aspose.Cells smart markers row level event handling C#
// Tags: WorkbookDesigner notify property Aspose.Cells | smart marker row insertion event C# | Aspose.Cells SmartMarkerEventHandler usage | enable row callbacks for smart markers | process smart markers with notifications

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads or creates an Excel template containing smart markers, sets WorkbookDesigner.Notify to true, registers a SmartMarkerEventHandler that is invoked for each row inserted during marker processing, binds a DataTable as the data source, processes the markers, and saves the populated workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Prepare input workbook (load if exists, otherwise create a simple template)
            string inputPath = "input.xlsx";
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Template";

                // Sample smart markers
                ws.Cells["A1"].PutValue("&=Employees.Name");
                ws.Cells["B1"].PutValue("&=Employees.Age");
                ws.Cells["C1"].PutValue("&=Employees.Department");
            }

            // Create a WorkbookDesigner to work with smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set the data source for the smart markers
            DataTable dt = GetSampleData();
            designer.SetDataSource("Employees", dt);

            // Process the smart markers
            designer.Process();

            // Save the resulting workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Sample method to generate a DataTable as a data source (replace with real data)
    private static DataTable GetSampleData()
    {
        var table = new DataTable("Employees");
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Age", typeof(int));
        table.Columns.Add("Department", typeof(string));

        table.Rows.Add("Alice", 30, "HR");
        table.Rows.Add("Bob", 35, "IT");
        table.Rows.Add("Charlie", 28, "Finance");

        return table;
    }
}
