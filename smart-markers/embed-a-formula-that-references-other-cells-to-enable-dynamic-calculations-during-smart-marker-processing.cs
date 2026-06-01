using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerFormulaDemo
{
    // Simple data class for demonstration
    public class Product
    {
        public string Name { get; set; } = string.Empty;   // avoid nullable warning
        public double Price { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Define smart markers in the template
                //    & =Products.Name  -> will be replaced with product name
                //    & =Products.Price -> will be replaced with product price
                cells["A2"].PutValue("&=Products.Name");
                cells["B2"].PutValue("&=Products.Price");

                // 3. Embed a formula that references the price column (dynamic calculation)
                //    Column C will show price with 20% tax: =B2*1.2
                //    Use overload that specifies the formula is not an array formula
                cells["C2"].SetFormula("=B2*1.2", false);

                // 4. Prepare sample data source
                List<Product> products = new List<Product>
                {
                    new Product { Name = "Apple",  Price = 1.00 },
                    new Product { Name = "Banana", Price = 0.50 },
                    new Product { Name = "Cherry", Price = 2.00 }
                };

                // 5. Set up WorkbookDesigner, assign data source, enable formula calculation
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("Products", products);
                designer.CalculateFormula = true; // calculate formulas after smart marker processing

                // 6. Process smart markers (rule: Process())
                designer.Process();

                // 7. Save the resulting workbook (lifecycle rule: save)
                string outputPath = "SmartMarkerWithFormula.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}