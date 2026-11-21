// Title: Using Aspose.Cells WorkbookDesigner to load a multi‑sheet template and apply separate smart‑marker data sources per worksheet in C#
// Common Searches: Aspose.Cells C# load workbook with multiple template sheets and process smart markers | How to assign different data sources to each worksheet using WorkbookDesigner | Processing smart markers on separate sheets in the same workbook with Aspose.Cells | C# example for multi‑template Excel using WorkbookDesigner and smart markers | Set data source for People and Products sheets in Aspose.Cells WorkbookDesigner
// Tags: WorkbookDesigner load multiple worksheets | smart markers per sheet | assign distinct data sources Aspose.Cells | process multi‑template Excel workbook | C# Aspose.Cells template sheets

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMultipleTemplates
{
    // Sample data classes
    // The sample loads a multi‑sheet Excel template (MultiTemplate.xlsx) into a Workbook, creates a WorkbookDesigner, and assigns two separate data sources—a list of Person objects named "People" and a list of Product objects named "Products". Calling Process() populates the smart markers on each corresponding worksheet, and the result is saved as Result.xlsx. The code includes basic error handling for missing files and processing exceptions.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Product
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public double Price { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            const string templatePath = "MultiTemplate.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file \"{templatePath}\" not found.");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook that contains multiple template sheets
                workbook = new Workbook(templatePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load template workbook: {ex.Message}");
                return;
            }

            try
            {
                var designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Prepare data sources for each sheet
                var people = new List<Person>
                {
                    new Person { Name = "John", Age = 30 },
                    new Person { Name = "Alice", Age = 25 }
                };

                var products = new List<Product>
                {
                    new Product { Id = 1, Name = "Apple", Price = 1.20 },
                    new Product { Id = 2, Name = "Banana", Price = 0.80 }
                };

                // Set data sources
                designer.SetDataSource("People", people);
                designer.SetDataSource("Products", products);

                // Process all worksheets (markers will be applied to the appropriate sheets)
                designer.Process();

                // Save the resulting workbook
                designer.Workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to \"{resultPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }
    }
}
