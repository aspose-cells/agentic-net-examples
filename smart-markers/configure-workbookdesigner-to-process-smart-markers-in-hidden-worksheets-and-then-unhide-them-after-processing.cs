// Title: How to process smart markers in hidden worksheets with WorkbookDesigner and restore original sheet visibility using Aspose.Cells for .NET
// AI Prompts: Generate C# code that temporarily sets all worksheets to visible, runs WorkbookDesigner.Process to evaluate smart markers, and then reverts each sheet to its original IsVisible state. | Show how to bind a DataTable to a smart‑marker name and save the workbook after processing while keeping hidden sheets hidden. | Provide a robust example that loads a template, handles a missing file fallback, and preserves worksheet visibility when using Aspose.Cells smart markers.
// Common Searches: How to evaluate smart markers in worksheets that are hidden in an Aspose.Cells workbook | Make hidden worksheets visible only during WorkbookDesigner smart marker processing in C# | Restore hidden sheet status after smart marker replacement with Aspose.Cells | Binding a DataTable to a smart marker located on a hidden worksheet | C# example for processing smart markers while preserving original sheet visibility
// Tags: smart marker processing on hidden sheets Aspose.Cells | WorkbookDesigner handling of sheet visibility during processing | revert worksheet visibility after WorkbookDesigner processing | DataTable as data source for smart markers Aspose.Cells | export processed workbook to XLSX with Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerHiddenSheetDemo
{
    // The example loads a workbook (or creates a fallback), records each worksheet's IsVisible state, temporarily makes all sheets visible, binds a DataTable named Employees to the smart‑marker '&=Employees.Name', processes all smart markers with WorkbookDesigner, restores the original visibility of each worksheet, and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the template workbook that may contain hidden sheets with smart markers
                const string templatePath = "TemplateWithHiddenSheets.xlsx";

                // Load the workbook if it exists; otherwise create a simple workbook as a fallback
                Workbook workbook;
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    Console.WriteLine($"Template file '{templatePath}' not found. Creating a new workbook as fallback.");
                    workbook = new Workbook();
                    // Add a simple smart marker for demonstration purposes
                    workbook.Worksheets[0].Cells["A1"].PutValue("&=Employees.Name");
                }

                // Preserve original visibility states of all worksheets
                bool[] originalVisibility = new bool[workbook.Worksheets.Count];
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    originalVisibility[i] = workbook.Worksheets[i].IsVisible;
                    // Make every worksheet visible temporarily so smart markers are processed
                    workbook.Worksheets[i].IsVisible = true;
                }

                // Initialize WorkbookDesigner with the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Example data source – a simple DataTable
                DataTable dt = new DataTable("Employees");
                dt.Columns.Add("Name", typeof(string));
                dt.Columns.Add("Age", typeof(int));
                dt.Rows.Add("John Doe", 30);
                dt.Rows.Add("Jane Smith", 28);

                // Bind the data source to a name used in the smart markers
                designer.SetDataSource("Employees", dt);

                // Process all smart markers (including those that were originally in hidden sheets)
                designer.Process();

                // Restore the original visibility of each worksheet
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    workbook.Worksheets[i].IsVisible = originalVisibility[i];
                }

                // Save the processed workbook
                const string outputPath = "ProcessedOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
