// Title: Toggle Smart Marker sections with WorkbookDesigner.SetVariable Boolean flag in C# using Aspose.Cells
// AI Prompts: Write C# code that loads an Excel template, creates a WorkbookDesigner, calls SetVariable("ShowSection", true) to control a Smart Marker block, processes the markers, and saves the workbook. | Show how to replace a DataSet data source with WorkbookDesigner.SetVariable to conditionally display or hide Smart Marker sections based on a Boolean variable in Aspose.Cells. | Provide a step‑by‑step tutorial for using WorkbookDesigner.SetVariable to pass a true/false flag that drives Smart Marker visibility in a .NET Excel report.
// Common Searches: Aspose.Cells C# setvariable to hide smart marker block | How to conditionally display smart marker sections using WorkbookDesigner.SetVariable | Toggle visibility of Excel smart markers with a boolean variable in .NET | SetVariable method example for smart markers in Aspose.Cells C#
// Tags: workbookdesigner setvariable boolean | smart marker conditional visibility aspnet cells | c# toggle smart marker section | excel template smart markers setvariable | aspocells setvariable example

using Aspose.Cells;
using System;
using System.Data;
using System.IO;

// The example demonstrates loading an Excel template, creating a WorkbookDesigner, using SetVariable to pass a Boolean flag that determines whether a Smart Marker block is shown or hidden, processing the markers, and saving the resulting workbook, with proper error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "template.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: Template file \"{templatePath}\" not found.");
                return;
            }

            // Load the workbook that contains smart markers
            Workbook workbook = new Workbook(templatePath);

            // Create a WorkbookDesigner for processing smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare a DataSet as the data source for smart markers
            DataSet dataSet = new DataSet();
            DataTable table = new DataTable("Data");
            table.Columns.Add("ShowSection", typeof(bool));
            table.Rows.Add(true);
            dataSet.Tables.Add(table);

            // Assign the data source to the designer
            designer.SetDataSource(dataSet);

            // Process the smart markers
            designer.Process();

            // Save the resulting workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
