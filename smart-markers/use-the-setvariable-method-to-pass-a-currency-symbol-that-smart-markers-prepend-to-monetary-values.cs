// Title: C# example: prepend a currency symbol to smart marker values using a global variable with Aspose.Cells WorkbookDesigner
// AI Prompts: Generate C# code that defines a Globals class with a CurrencySymbol property, sets an instance as a data source for WorkbookDesigner, and uses ${Globals.CurrencySymbol}{CellReference} in smart markers to add the symbol to price cells. | Show how to process smart markers after defining a global currency variable and save the resulting workbook to a file with Aspose.Cells.
// Common Searches: how to add a custom currency symbol to smart marker output in Aspose.Cells C# | Aspose.Cells WorkbookDesigner global variable for currency formatting example | C# smart markers prepend symbol to numeric values using Aspose.Cells | using SetDataSource to pass variables to smart markers in Aspose.Cells | Aspose.Cells smart marker expression for currency symbol and price column
// Tags: WorkbookDesigner SetDataSource global variable | Aspose.Cells smart marker currency formatting | C# prepend symbol to smart marker values | Excel export with currency symbol using Aspose.Cells | smart marker variable substitution Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerExample
{
    // Simple class to hold global variables for smart markers
    // The sample creates a workbook, defines a Globals class containing a CurrencySymbol property, assigns an instance to WorkbookDesigner via SetDataSource, and references the symbol in smart marker cells using the ${Globals.CurrencySymbol}{B2} syntax. After processing the markers, the workbook is saved as SmartMarkerCurrency.xlsx.
    public class Globals
    {
        public string CurrencySymbol { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Initialize WorkbookDesigner for smart marker processing
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Set a global variable (currency symbol) via a data source object
                designer.SetDataSource("Globals", new Globals { CurrencySymbol = "$" });

                // Sample data that will be referenced by smart markers
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1.5);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(0.8);

                // Smart marker expressions that prepend the currency symbol to the price values
                sheet.Cells["C1"].PutValue("Price with Symbol");
                sheet.Cells["C2"].PutValue("${Globals.CurrencySymbol}{B2}");
                sheet.Cells["C3"].PutValue("${Globals.CurrencySymbol}{B3}");

                // Process the smart markers
                designer.Process();

                // Define output file path
                string outputPath = "SmartMarkerCurrency.xlsx";

                // Ensure the output directory exists
                string fullOutputPath = Path.GetFullPath(outputPath);
                string outputDir = Path.GetDirectoryName(fullOutputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(fullOutputPath);
                Console.WriteLine($"Workbook saved successfully to '{fullOutputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
