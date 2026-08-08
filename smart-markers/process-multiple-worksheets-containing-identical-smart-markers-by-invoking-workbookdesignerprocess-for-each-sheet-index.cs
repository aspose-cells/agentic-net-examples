// Title: Process identical Smart Markers on every worksheet using WorkbookDesigner.Process (sheet index) – Aspose.Cells for .NET
// Description: C# example that loads a workbook with the same smart markers on each sheet, binds a DataTable named "Employees" as the data source, loops through all worksheet indexes and calls WorkbookDesigner.Process(index, true) to fill the markers while keeping any unknown markers, then saves the result.
// Keywords: Aspose.Cells smart markers multiple sheets | WorkbookDesigner.Process sheet index | preserve unknown smart markers | bind DataTable to smart markers | C# Excel smart marker processing
// Common Searches: Aspose.Cells process smart markers on each worksheet | WorkbookDesigner.Process overload for specific sheet | keep unrecognized smart markers Aspose.Cells | iterate worksheets to fill smart markers C# | smart marker loop hidden sheets Aspose
// Developer Intent: Run WorkbookDesigner.Process for every worksheet index to populate identical smart markers across a workbook while optionally preserving markers that have no matching data.
// Use Cases: Create employee‑list reports where every sheet uses the same smart markers and shares one DataTable. | Populate a multi‑month sales template that repeats identical smart markers on each month’s worksheet. | Generate a multi‑sheet invoice workbook where customer and product smart markers appear on every page.
// AI Prompts: Show how to skip hidden worksheets when processing smart markers with WorkbookDesigner.Process. | Provide code to process smart markers only on worksheets whose names are in a specified list. | Explain error handling for individual sheet processing so the loop continues with remaining sheets.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerProcessing
{
    // C# example that loads a workbook with the same smart markers on each sheet, binds a DataTable named "Employees" as the data source, loops through all worksheet indexes and calls WorkbookDesigner.Process(index, true) to fill the markers while keeping any unknown markers, then saves the result.
    public class MultipleSheetProcessor
    {
        public static void Run()
        {
            try
            {
                const string templatePath = "TemplateWithSmartMarkers.xlsx";
                const string outputPath = "ProcessedMultipleSheets.xlsx";

                // Verify that the template file exists to avoid FileNotFoundException
                if (!File.Exists(templatePath))
                {
                    throw new FileNotFoundException($"Template file not found: {templatePath}");
                }

                // Load the workbook that contains identical smart markers on each worksheet
                Workbook workbook = new Workbook(templatePath);

                // Prepare a simple data source (DataTable) that matches the smart markers
                DataTable data = new DataTable("Employees");
                data.Columns.Add("Name", typeof(string));
                data.Columns.Add("Age", typeof(int));
                data.Rows.Add("John Doe", 30);
                data.Rows.Add("Jane Smith", 28);
                data.Rows.Add("Bob Johnson", 45);

                // Initialize WorkbookDesigner with the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Bind the data source to a name used in the smart markers (e.g., "Employees")
                designer.SetDataSource("Employees", data);

                // Process each worksheet individually using the sheet index overload
                // The boolean parameter 'true' preserves any unrecognized smart markers
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    designer.Process(i, true);
                }

                // Save the processed workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            MultipleSheetProcessor.Run();
        }
    }
}
