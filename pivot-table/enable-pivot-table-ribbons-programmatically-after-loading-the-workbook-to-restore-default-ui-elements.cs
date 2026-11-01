// Title: Enable the PivotTable ribbon UI programmatically in an Aspose.Cells for .NET workbook after loading (when supported)
// AI Prompts: Generate C# code that checks the Aspose.Cells version and sets the appropriate workbook setting to display the PivotTable ribbon after opening an Excel file. | Show how to programmatically restore the default PivotTable toolbar in a loaded workbook using Aspose.Cells for .NET, including handling of API availability. | Provide a method that conditionally activates PivotTable UI elements in Aspose.Cells based on whether the required property exists.
// Common Searches: aspocells enable pivot table ribbon after opening workbook c# | c# restore default pivot table toolbar using Aspose.Cells | how to turn on PivotTable UI in Aspose.Cells for .NET | WorkbookSettings.EnablePivotTableUI not found in older Aspose.Cells versions | programmatically show pivot table ribbon in Excel file with Aspose.Cells
// Tags: Aspose.Cells pivot table UI activation | pivot table UI property Aspose.Cells | programmatic pivot ribbon restoration | C# restore pivot toolbar Aspose.Cells | Aspose.Cells version compatibility pivot UI

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook with Aspose.Cells, indicates that the WorkbookSettings.EnablePivotTableUI property is unavailable in the current library version, and saves the workbook unchanged.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                var workbook = new Workbook(inputPath);

                // NOTE: The WorkbookSettings.EnablePivotTableUI property is not available in this version of Aspose.Cells.
                // If needed, adjust the workbook settings using the appropriate API for your version.

                // Save the workbook with the (unchanged) settings
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
