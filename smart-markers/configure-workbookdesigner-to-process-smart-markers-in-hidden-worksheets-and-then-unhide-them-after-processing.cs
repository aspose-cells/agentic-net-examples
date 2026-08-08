// Title: C# – Process Smart Markers in Hidden Worksheets with WorkbookDesigner
// Description: Load an Excel workbook, temporarily unhide hidden sheets, bind a DataTable to WorkbookDesigner, process all smart markers (including those on hidden sheets), then restore the original visibility before saving. Demonstrates Aspose.Cells handling of hidden worksheets in C#.
// Keywords: Aspose.Cells | WorkbookDesigner | smart markers | hidden worksheets | C# | .NET | Excel template processing | temporary unhide | restore sheet visibility | DataTable data source | US developers | European .NET community
// Common Searches: Aspose.Cells process smart markers on hidden sheets | C# unhide hidden worksheets for WorkbookDesigner | how to keep sheets hidden after smart marker processing | Aspose.Cells temporary sheet visibility C# example | process hidden worksheet smart markers Aspose.Cells
// Developer Intent: Fill smart markers on all sheets, including hidden ones, while preserving the workbook’s original hidden‑sheet layout.
// Use Cases: Generate multi‑sheet reports where template sheets with smart markers stay hidden from end users. | Automate confidential calculations in hidden worksheets, populate them programmatically, and deliver a clean workbook. | Create Excel templates that use hidden placeholder sheets for data merging, then unhide, process, and re‑hide them in a single workflow.
// AI Prompts: Write C# code using Aspose.Cells to unhide hidden worksheets, run WorkbookDesigner.Process on smart markers, and restore the original hidden state. | Provide a reusable method that accepts a Workbook and a DataTable, processes smart markers on every sheet (including hidden ones), and ensures hidden sheets remain hidden after processing. | Explain step‑by‑step how to detect hidden worksheets, make them visible for smart marker processing, and revert their visibility using Aspose.Cells APIs.

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

// Load an Excel workbook, temporarily unhide hidden sheets, bind a DataTable to WorkbookDesigner, process all smart markers (including those on hidden sheets), then restore the original visibility before saving. Demonstrates Aspose.Cells handling of hidden worksheets in C#.
class ProcessHiddenSmartMarkers
{
    static void Main()
    {
        try
        {
            const string templatePath = "TemplateWithHiddenSheets.xlsx";

            // Load existing template or create a new workbook if the file is missing.
            Workbook workbook;
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                // Create a simple workbook with a visible sheet and a hidden sheet containing smart markers.
                workbook = new Workbook();

                // Visible sheet with sample data.
                Worksheet visibleWs = workbook.Worksheets[0];
                visibleWs.Name = "VisibleSheet";
                visibleWs.Cells["A1"].PutValue("Name");
                visibleWs.Cells["B1"].PutValue("Value");
                visibleWs.Cells["A2"].PutValue("&=Sample.Name");
                visibleWs.Cells["B2"].PutValue("&=Sample.Value");

                // Hidden sheet with smart markers.
                Worksheet hiddenWs = workbook.Worksheets.Add("HiddenSheet");
                hiddenWs.IsVisible = false;
                hiddenWs.Cells["A1"].PutValue("Hidden Name");
                hiddenWs.Cells["B1"].PutValue("Hidden Value");
                hiddenWs.Cells["A2"].PutValue("&=Sample.Name");
                hiddenWs.Cells["B2"].PutValue("&=Sample.Value");
            }

            // Remember which worksheets were originally hidden.
            List<int> originallyHidden = new List<int>();
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                if (!workbook.Worksheets[i].IsVisible)
                {
                    originallyHidden.Add(i);
                    // Unhide temporarily to allow WorkbookDesigner to process its smart markers.
                    workbook.Worksheets[i].SetVisible(true, true);
                }
            }

            // Initialize WorkbookDesigner with the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set up a data source for the smart markers.
            DataTable dt = new DataTable("Sample");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Value", typeof(double));
            dt.Rows.Add("Item1", 10.0);
            dt.Rows.Add("Item2", 20.0);
            designer.SetDataSource(dt);

            // Process all smart markers, including those that were originally in hidden sheets.
            designer.Process();

            // Restore the original hidden state of the worksheets.
            foreach (int idx in originallyHidden)
            {
                workbook.Worksheets[idx].SetVisible(false, true);
            }

            // Save the processed workbook.
            const string outputPath = "ProcessedOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
