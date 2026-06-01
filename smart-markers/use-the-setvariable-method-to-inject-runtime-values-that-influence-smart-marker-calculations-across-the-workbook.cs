using System;
using System.IO;
using Aspose.Cells; // WorkbookDesigner is in this namespace

namespace AsposeCellsVariableDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Add a worksheet that will hold variable definitions
                Worksheet varSheet = workbook.Worksheets.Add("Variables");
                // Put variable names in column A and placeholder values in column B
                varSheet.Cells["A1"].PutValue("DiscountRate");
                varSheet.Cells["B1"].PutValue(0.0); // initial placeholder

                varSheet.Cells["A2"].PutValue("TaxRate");
                varSheet.Cells["B2"].PutValue(0.0); // initial placeholder

                // Inject runtime values directly into the variable sheet
                varSheet.Cells["B1"].PutValue(0.15); // 15% discount
                varSheet.Cells["B2"].PutValue(0.08); // 8% tax

                // Create a WorkbookDesigner and associate it with the workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook)
                {
                    // Tell the designer which worksheet contains the variables
                    VariablesWorksheetName = "Variables"
                };

                // Add a template worksheet that uses the variables via smart markers
                Worksheet template = workbook.Worksheets.Add("Template");
                // Smart markers that reference the variables
                template.Cells["A1"].PutValue("Discount Rate:");
                template.Cells["B1"].PutValue("&=$DiscountRate"); // variable smart marker

                template.Cells["A2"].PutValue("Tax Rate:");
                template.Cells["B2"].PutValue("&=$TaxRate"); // variable smart marker

                // Process the smart markers – they will be replaced with the injected values
                designer.Process();

                // Save the workbook (lifecycle rule: save)
                string outputPath = "VariableInjectionDemo.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with injected variable values to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}