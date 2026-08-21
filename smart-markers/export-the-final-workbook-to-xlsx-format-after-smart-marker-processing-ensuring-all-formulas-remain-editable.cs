// Title: Export Smart Marker Workbook to XLSX with Editable Formulas – Aspose.Cells for .NET
// Description: C# code that loads a template workbook with Aspose.Cells smart markers, binds a custom data source, processes the markers while preserving unknown tags, disables automatic calculation on save, and saves the file as XLSX so all formulas remain editable for downstream editing.
// Keywords: Aspose.Cells C# smart markers export | save workbook as xlsx without recalculating formulas | WorkbookDesigner CalculateOnSave false | editable formulas after smart marker processing | .NET spreadsheet automation | preserve formulas Aspose.Cells | smart marker data binding C# | export template to xlsx Aspose | global spreadsheet generation
// Common Searches: how to export smart marker workbook to xlsx without formula recalculation | aspnet keep formulas editable after processing smart markers | disable calculate on save Aspose.Cells C# | save processed smart markers as xlsx preserving formulas | Aspose.Cells smart marker export example
// Developer Intent: Save a workbook that has been processed with smart markers as an XLSX file while preventing automatic formula recalculation.
// Use Cases: Generate personalized reports from a template and let users adjust totals or other calculations after the data merge. | Create invoices where tax and discount formulas stay editable for accountants to modify post‑generation. | Export data‑driven spreadsheets for downstream analytics while retaining all original calculation logic.
// AI Prompts: Write C# code using Aspose.Cells to process smart markers in a template and save the result as XLSX without triggering formula recalculation. | Explain the impact of Workbook.Settings.FormulaSettings.CalculateOnSave on formula editability after smart marker processing. | Show how to bind a List of custom objects to a smart marker data source and preserve all formulas when saving the workbook.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// C# code that loads a template workbook with Aspose.Cells smart markers, binds a custom data source, processes the markers while preserving unknown tags, disables automatic calculation on save, and saves the file as XLSX so all formulas remain editable for downstream editing.
public class SmartMarkerExport
{
    public static void Run()
    {
        try
        {
            string templatePath = "template.xlsx";
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load the template workbook that contains smart markers
            Workbook workbook = new Workbook(templatePath);

            // Create a WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Prepare a sample data source (replace with your actual data)
            ArrayList persons = new ArrayList
            {
                new Person { Name = "John Doe", Age = 30 },
                new Person { Name = "Jane Smith", Age = 28 }
            };

            // Bind the data source to a name used in the smart markers
            designer.SetDataSource("Persons", persons);

            // Process the smart markers; true = preserve unrecognized markers
            designer.Process(true);

            // Ensure formulas stay editable (do not recalculate on save)
            workbook.Settings.FormulaSettings.CalculateOnSave = false;

            // Save the processed workbook to XLSX format
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Sample data class used in the data source
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }
}

// Entry point
public class Program
{
    public static void Main()
    {
        SmartMarkerExport.Run();
    }
}
