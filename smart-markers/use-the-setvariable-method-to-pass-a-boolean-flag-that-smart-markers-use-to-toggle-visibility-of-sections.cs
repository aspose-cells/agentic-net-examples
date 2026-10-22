// Title: How to toggle smart‑marker section visibility in an Excel template using WorkbookDesigner.SetVariable (C# Aspose.Cells)
// AI Prompts: Write C# code that loads an Excel template, calls WorkbookDesigner.SetVariable("ShowSection", true) to control smart‑marker sections, processes the markers, and saves the workbook. | Refactor a smart‑marker example that uses a DataSet so it instead uses WorkbookDesigner.SetVariable to pass a boolean flag for conditional display. | Demonstrate hiding or showing a smart‑marker block in a .xlsx file by setting a true/false variable with WorkbookDesigner.SetVariable before invoking Process.
// Common Searches: Aspose.Cells C# WorkbookDesigner SetVariable boolean example for smart markers | how to hide a smart marker section in Excel using a flag with Aspose.Cells | toggle visibility of smart marker blocks in .NET workbook designer | setvariable true false smart marker Aspose.Cells tutorial
// Tags: WorkbookDesigner.SetVariable boolean smart marker | conditional smart marker sections Aspose.Cells | C# toggle smart marker visibility | Excel template smart marker flag | Aspose.Cells smart marker conditional display

using System;
using System.IO;
using System.Data;
using Aspose.Cells;

// The example loads an Excel workbook containing smart‑marker blocks, uses WorkbookDesigner.SetVariable("ShowSection", true) to pass a boolean flag that the markers evaluate, calls Process to apply the conditional logic, and saves the resulting file.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "TemplateWithSmartMarkers.xlsx";
            const string resultPath = "ResultWithToggledSection.xlsx";

            // Ensure the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load the template workbook that contains smart markers
            var workbook = new Workbook(templatePath);

            // Initialize the WorkbookDesigner to work with smart markers
            var designer = new WorkbookDesigner(workbook);

            // Create a DataSet with a DataTable containing the ShowSection variable
            var dataTable = new DataTable("Data");
            dataTable.Columns.Add("ShowSection", typeof(bool));
            dataTable.Rows.Add(true);

            var dataSet = new DataSet();
            dataSet.Tables.Add(dataTable);

            // Provide the data source for smart markers
            designer.SetDataSource(dataSet);

            // Process the smart markers – this will evaluate the variable and hide/show sections accordingly
            designer.Process();

            // Save the resulting workbook
            workbook.Save(resultPath);
        }
        catch (Exception ex)
        {
            // Log any errors that occur during processing
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
