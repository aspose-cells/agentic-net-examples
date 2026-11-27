// Title: Inject runtime values into smart markers with WorkbookDesigner.SetDataSource in Aspose.Cells for .NET
// Common Searches: how to set smart marker values at runtime using WorkbookDesigner in C# | Aspose.Cells example for assigning TotalSales and ReportDate to smart markers | C# code to process smart markers after setting custom variables with SetDataSource | load smart marker template, set data source values, and save workbook using Aspose.Cells | runtime injection of region and sales data into Excel smart markers with Aspose.Cells
// Tags: WorkbookDesigner SetDataSource smart markers | runtime data injection Aspose.Cells | C# smart marker variable assignment | process smart markers Excel workbook | calculate formulas after smart marker processing

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a SmartMarkerTemplate.xlsx workbook, uses WorkbookDesigner.SetDataSource to assign values for TotalSales, ReportDate, and Region, processes the smart markers, recalculates formulas, and saves the result to SmartMarkerResult.xlsx while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        const string templatePath = "SmartMarkerTemplate.xlsx";
        const string resultPath = "SmartMarkerResult.xlsx";

        try
        {
            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the workbook that contains smart markers
            var workbook = new Workbook(templatePath);

            // Use WorkbookDesigner to set smart marker variables
            var designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("TotalSales", 12500.75);
            designer.SetDataSource("ReportDate", DateTime.Now);
            designer.SetDataSource("Region", "North America");

            // Process the smart markers
            designer.Process();

            // Optional: recalculate formulas after processing
            workbook.CalculateFormula();

            // Save the resulting workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to {resultPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
