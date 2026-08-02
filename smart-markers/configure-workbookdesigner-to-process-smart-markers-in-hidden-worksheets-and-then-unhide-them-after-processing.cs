// Title: Process Smart Markers in Hidden Sheets with WorkbookDesigner (C#) and Reveal All Worksheets
// Description: Load or create a workbook that contains hidden worksheets with smart markers, bind a DataTable, run WorkbookDesigner.Process to merge data, then set each sheet’s IsVisible to true and save the file.
// Keywords: Aspose.Cells | WorkbookDesigner | smart markers | hidden worksheets | unhide sheets | C# | .NET | data binding | Excel template | report generation
// Common Searches: Aspose.Cells process smart markers on hidden sheet | WorkbookDesigner hidden worksheet visibility | C# unhide Excel sheets after smart marker processing | include hidden sheets in Aspose.Cells smart marker merge | generate report with hidden template sheets Aspose
// Developer Intent: Merge data into smart markers placed on hidden worksheets and make every sheet visible before saving the workbook.
// Use Cases: Create a multi‑sheet report where template sheets are hidden during data merge and revealed for the final output. | Build an invoice workbook that stores calculation templates on hidden sheets, then unhide them after populating smart markers for client review. | Generate a dashboard workbook that uses hidden sheets for intermediate smart‑marker data, processing them with WorkbookDesigner and exposing all sheets in the published file.
// AI Prompts: Write C# code using Aspose.Cells to process smart markers in hidden worksheets and then unhide all sheets before saving. | Explain how WorkbookDesigner handles smart markers on hidden sheets and what steps are required to include them in the final workbook. | Provide a step‑by‑step guide to create an Excel template with hidden smart‑marker sheets, bind a DataTable, process the markers, and make the sheets visible.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// Load or create a workbook that contains hidden worksheets with smart markers, bind a DataTable, run WorkbookDesigner.Process to merge data, then set each sheet’s IsVisible to true and save the file.
class ProcessHiddenSmartMarkers
{
    static void Main()
    {
        try
        {
            string templatePath = "TemplateWithHiddenSheets.xlsx";
            Workbook workbook;

            // Load existing template or create a minimal one if it does not exist.
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                // Create a workbook with a visible sheet and a hidden sheet containing smart markers.
                workbook = new Workbook();
                Worksheet visibleWs = workbook.Worksheets[0];
                visibleWs.Name = "VisibleSheet";
                visibleWs.Cells["A1"].PutValue("&=Employees.Name");
                visibleWs.Cells["B1"].PutValue("&=Employees.Age");

                Worksheet hiddenWs = workbook.Worksheets.Add("HiddenSheet");
                hiddenWs.IsVisible = false;
                hiddenWs.Cells["A1"].PutValue("&=Employees.Name");
                hiddenWs.Cells["B1"].PutValue("&=Employees.Age");
            }

            // Initialize the designer with the workbook.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare sample data matching the smart markers.
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("John Doe", 30);
            dt.Rows.Add("Jane Smith", 28);

            // Bind the data source.
            designer.SetDataSource("Employees", dt);

            // Process all smart markers, including those in hidden sheets.
            designer.Process();

            // Unhide all worksheets after processing.
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.IsVisible = true;
            }

            // Save the processed workbook.
            string outputPath = "ProcessedOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
