// Title: Combine multiple Excel templates into a single WorkbookDesigner and assign distinct smart‑marker data sources per worksheet (Aspose.Cells for .NET)
// AI Prompts: Load two Excel template files, copy all their worksheets into one WorkbookDesigner instance, bind a List<Person> to Sheet1 and a List<Product> to Sheet2, process the smart markers, and save the merged workbook. | Merge any number of smart‑marker templates by copying their sheets into a shared WorkbookDesigner, set a different POCO collection as the data source for each sheet, run Process(), and output the combined result.
// Common Searches: how to merge several Excel templates with different smart marker data sources using Aspose.Cells | Aspose.Cells set separate data source for each worksheet in WorkbookDesigner | combine multiple smart marker templates into one workbook .NET | copy worksheets from multiple workbooks into a single WorkbookDesigner Aspose.Cells | process smart markers after merging Excel templates Aspose.Cells
// Tags: merge worksheets WorkbookDesigner Aspose.Cells | worksheet‑specific smart marker data source | process smart markers combined workbook | copy template sheets between workbooks | use POCO collections with smart markers Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Shows how to load two Excel template files, copy their worksheets into a single WorkbookDesigner, assign List<Person> to Sheet1 and List<Product> to Sheet2 as separate smart‑marker data sources, process all markers, and save the merged result as MergedResult.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Verify template files exist to avoid FileNotFoundException
            const string templatePath1 = "Template1.xlsx";
            const string templatePath2 = "Template2.xlsx";

            if (!File.Exists(templatePath1))
                throw new FileNotFoundException($"Template file not found: {templatePath1}");
            if (!File.Exists(templatePath2))
                throw new FileNotFoundException($"Template file not found: {templatePath2}");

            // Create a WorkbookDesigner instance
            WorkbookDesigner designer = new WorkbookDesigner
            {
                // Initialize an empty workbook that will hold all templates
                Workbook = new Workbook()
            };

            // ---------- Load first template ----------
            Workbook template1 = new Workbook(templatePath1);
            // Copy all worksheets from the first template into the designer's workbook
            foreach (Worksheet ws in template1.Worksheets)
            {
                designer.Workbook.Worksheets.AddCopy(ws.Name);
            }

            // ---------- Load second template ----------
            Workbook template2 = new Workbook(templatePath2);
            // Copy all worksheets from the second template into the designer's workbook
            foreach (Worksheet ws in template2.Worksheets)
            {
                designer.Workbook.Worksheets.AddCopy(ws.Name);
            }

            // ---------- Prepare data sources ----------
            // Data for the first sheet (assumed to be named "Sheet1" in Template1.xlsx)
            List<Person> persons = new List<Person>
            {
                new Person { Name = "John", Age = 30 },
                new Person { Name = "Jane", Age = 25 }
            };

            // Data for the second sheet (assumed to be named "Sheet2" in Template2.xlsx)
            List<Product> products = new List<Product>
            {
                new Product { Id = 1, Name = "Apple",  Price = 0.5 },
                new Product { Id = 2, Name = "Banana", Price = 0.3 }
            };

            // ---------- Set data sources for each sheet ----------
            designer.SetDataSource("Sheet1", persons);
            designer.SetDataSource("Sheet2", products);

            // Process all smart markers in the combined workbook
            designer.Process();

            // Save the final workbook containing both processed sheets
            designer.Workbook.Save("MergedResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Sample data class for the first sheet
    public class Person
    {
        public string Name { get; set; } = null!;
        public int Age { get; set; }
    }

    // Sample data class for the second sheet
    public class Product
    {
        public int Id { get; set; }
        public string Name { get; set; } = null!;
        public double Price { get; set; }
    }
}
