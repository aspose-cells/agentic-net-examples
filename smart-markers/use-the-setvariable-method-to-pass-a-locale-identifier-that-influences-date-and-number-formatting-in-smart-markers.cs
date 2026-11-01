// Title: Use WorkbookDesigner.SetVariable to apply US locale (LCID 1033) for date and number formatting in Aspose.Cells smart markers (C#)
// AI Prompts: Generate C# code that calls WorkbookDesigner.SetVariable("LocaleId", 1033) and updates smart marker expressions to format dates and numbers using the US locale. | Refactor an existing Aspose.Cells smart‑marker example so the locale is supplied via SetVariable instead of Workbook.Settings.CultureInfo. | Demonstrate how to embed the locale variable in a smart‑marker placeholder (e.g., ${Date:LocaleId}) to achieve culture‑aware formatting.
// Common Searches: aspocells setvariable localeid for smart marker formatting c# | how to change smart marker culture using setvariable in aspocells | apply us locale to smart markers without modifying workbook settings | c# smart marker date number format based on lcid aspocells | using setvariable to control number formatting in excel template aspocells
// Tags: WorkbookDesigner.SetVariable locale | smart marker culture formatting Aspose.Cells | C# set LCID for smart markers | locale‑aware smart markers Excel template | Aspose.Cells US locale smart markers

using System;
using System.Data;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example loads a template workbook, defines the US locale identifier (LCID 1033) via WorkbookDesigner.SetVariable, provides a DataTable with Date and Amount columns, processes smart markers that reference the locale variable for culture‑specific date and number formatting, and saves the resulting workbook.
class SmartMarkerLocaleExample
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: Template file '{templatePath}' was not found.");
                return;
            }

            // Load the workbook that contains smart markers
            Workbook workbook = new Workbook(templatePath);

            // Set workbook culture (locale) to English - United States (LCID 1033)
            workbook.Settings.CultureInfo = new CultureInfo(1033);

            // Create a WorkbookDesigner for processing smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare a data source for the smart markers
            DataTable data = new DataTable();
            data.Columns.Add("Date", typeof(DateTime));
            data.Columns.Add("Amount", typeof(double));
            data.Rows.Add(DateTime.Now, 1234.56);
            data.Rows.Add(DateTime.Now.AddDays(1), 7890.12);

            // Assign the data source to the designer
            designer.SetDataSource(data);

            // Process the smart markers using the supplied data and locale
            designer.Process();

            // Save the resulting workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
