// Title: Process Identical Smart Markers on All Worksheets with WorkbookDesigner.Process (C# Aspose.Cells)
// Description: Loads or creates an Excel template containing smart markers, builds a DataTable named Employees, binds it to WorkbookDesigner, then iterates through every worksheet calling designer.Process(sheetIndex, true) to replace the markers while preserving any unknown placeholders, and finally saves the result.
// Keywords: Aspose.Cells | C# | WorkbookDesigner | smart markers | multiple worksheets | process | preserve unknown markers | DataTable binding | Excel template | loop worksheets
// Common Searches: Aspose.Cells process smart markers on each sheet C# | WorkbookDesigner.Process multiple worksheets example | preserve unknown smart markers Aspose.Cells | loop through worksheets to apply smart markers | create Excel template with smart markers programmatically
// Developer Intent: Execute WorkbookDesigner.Process for every worksheet to replace identical smart markers with bound data while leaving any unmatched markers untouched.
// Use Cases: Generate a multi‑sheet report where the same employee table appears on each page. | Apply a common smart‑marker template to several departmental worksheets without altering custom placeholders. | Preserve user‑defined placeholder text that is not linked to a data source during batch processing.
// AI Prompts: Show how to skip worksheets that contain no smart markers before calling Process. | Provide code to add a new worksheet with the same smart markers and then process all sheets. | Explain how to catch and handle missing data‑source errors when processing smart markers. | Demonstrate changing the smart‑marker delimiters for a specific workbook. | Convert the processed workbook to PDF after all sheets have been processed.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// Loads or creates an Excel template containing smart markers, builds a DataTable named Employees, binds it to WorkbookDesigner, then iterates through every worksheet calling designer.Process(sheetIndex, true) to replace the markers while preserving any unknown placeholders, and finally saves the result.
class ProcessMultipleSheets
{
    static void Main()
    {
        try
        {
            const string templatePath = "SmartMarkerTemplate.xlsx";
            Workbook workbook;

            // Load existing template or create a minimal one if it does not exist
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Sheet1";

                // Sample smart markers that match the data source
                ws.Cells["A1"].PutValue("&=Employees.Name");
                ws.Cells["B1"].PutValue("&=Employees.Age");

                // Save the generated template for subsequent runs
                workbook.Save(templatePath);
            }

            // Prepare a sample data source (DataTable) that matches the smart markers
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("John Doe", 30);
            dt.Rows.Add("Jane Smith", 28);

            // Initialize the WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Bind the data source to a name used in the smart markers
            designer.SetDataSource("Employees", dt);

            // Process smart markers on each worksheet
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Preserve any unrecognized smart markers (second parameter = true)
                designer.Process(i, true);
            }

            // Save the processed workbook
            workbook.Save("ProcessedOutput.xlsx");
            Console.WriteLine("Processing completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
