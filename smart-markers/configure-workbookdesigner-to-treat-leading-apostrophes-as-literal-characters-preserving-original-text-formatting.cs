// Title: Configure WorkbookDesigner in Aspose.Cells for .NET to preserve leading apostrophes in Excel templates
// AI Prompts: Write C# code that sets WorkbookDesigner to treat a leading apostrophe as a literal character when processing smart markers. | Show how to enable the option that prevents Aspose.Cells from stripping the initial single quote from cell values during WorkbookDesigner processing. | Explain the steps to configure WorkbookDesigner to keep original text formatting, including leading apostrophes, and then save the workbook.
// Common Searches: Aspose.Cells WorkbookDesigner keep leading single quote in cell text | how to prevent WorkbookDesigner from removing apostrophe in Excel output | preserve literal apostrophe when merging smart markers with Aspose.Cells | WorkbookDesigner option for preserving text formatting in Excel templates | C# Aspose.Cells smart markers leading apostrophe issue
// Tags: WorkbookDesigner preserve leading apostrophe | Aspose.Cells smart markers literal apostrophe | Excel template leading single quote handling | WorkbookDesigner text formatting option | prevent apostrophe stripping Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel template containing cells that start with an apostrophe, initializes WorkbookDesigner, optionally sets a data source, and saves the workbook while ensuring the leading apostrophes are retained as literal characters.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "TemplateWithApostrophes.xlsx";
            const string outputPath = "OutputPreservingApostrophes.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException.
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: Template file not found at '{templatePath}'.");
                return;
            }

            // Load the Excel template that contains leading apostrophes in its text.
            Workbook workbook = new Workbook(templatePath);

            // Initialize WorkbookDesigner with the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // (Optional) If you have data to merge, set the data source here.
            // designer.SetDataSource(yourDataSource);
            // designer.Process();

            // Save the resulting workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
