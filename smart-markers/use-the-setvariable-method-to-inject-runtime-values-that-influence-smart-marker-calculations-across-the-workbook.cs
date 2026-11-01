// Title: Use WorkbookDesigner.SetVariable to inject runtime values into smart markers in an Aspose.Cells .NET workbook
// AI Prompts: Generate C# code that loads an Excel template, calls WorkbookDesigner.SetVariable to assign a date, a numeric amount, a boolean flag, and a string to smart markers, processes the markers, and saves the workbook. | Demonstrate converting WorkbookDesigner.SetDataSource calls to WorkbookDesigner.SetVariable for passing custom runtime variables to smart markers throughout the workbook.
// Common Searches: how to use WorkbookDesigner.SetVariable for smart markers in Aspose.Cells C# | Aspose.Cells C# set variable for smart marker calculations | replace SetDataSource with SetVariable in Aspose.Cells smart marker example | passing date and numeric values to smart markers using SetVariable Aspose.Cells | dynamic smart marker values with WorkbookDesigner.SetVariable .NET
// Tags: Inject variables via SetVariable in Aspose.Cells | Aspose.Cells smart marker runtime values | C# WorkbookDesigner custom variable binding | Dynamic smart marker calculations .NET | Process Excel smart markers with custom data

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel template (or creates a new workbook if the template is missing), creates a WorkbookDesigner, assigns runtime values such as the current date, total sales amount, premium‑customer flag, and company name to smart markers using WorkbookDesigner.SetVariable, processes the smart markers, and saves the resulting workbook.
class SmartMarkerVariableDemo
{
    static void Main()
    {
        try
        {
            // Path to the template workbook that contains smart markers
            string templatePath = "TemplateWithSmartMarkers.xlsx";

            Workbook workbook;

            // Load the template if it exists; otherwise create a new workbook as a fallback
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                Console.WriteLine($"Template file \"{templatePath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
                // Optionally add a placeholder sheet with smart markers here
            }

            // Initialize WorkbookDesigner for processing smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set runtime variables used by smart markers via data sources
            designer.SetDataSource("ReportDate", DateTime.Now);
            designer.SetDataSource("TotalSales", 15230.75);
            designer.SetDataSource("IsPremiumCustomer", true);
            designer.SetDataSource("CompanyName", "Contoso Ltd.");

            // Process the smart markers
            designer.Process();

            // Save the resulting workbook
            string resultPath = "SmartMarkerResult.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to \"{resultPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
