// Title: Enable CheckMergedCells to import custom objects while preserving merged ranges in Aspose.Cells for .NET
// AI Prompts: Configure ImportTableOptions with CheckMergedCells = true and use Cells.ImportCustomObjects to load a List of custom objects into a worksheet. | Create a merged cell block, set InsertRows and CheckMergedCells in ImportTableOptions, then import data starting at a specified cell address. | After importing custom objects with merged‑cell handling, save the workbook to an XLSX file using Aspose.Cells.
// Common Searches: Aspose.Cells preserve merged cells when importing a list of objects in C# | ImportCustomObjects with CheckMergedCells enabled example .NET | How to keep merged range values while adding custom object data to Excel using Aspose.Cells | C# import custom objects into worksheet and retain merged ranges
// Tags: ImportCustomObjects merged cells | CheckMergedCells option | ImportTableOptions InsertRows | Aspose.Cells custom object import C# | merged range handling Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a merged range, enables CheckMergedCells in ImportTableOptions, imports a list of Customer objects into the worksheet starting at A1, and saves the workbook as MergedImportDemo.xlsx.
    public class ImportCustomObjectsWithMergedCellsDemo
    {
        // Sample custom object
        public class Customer
        {
            public int CustomerId { get; set; }
            public string Name { get; set; } = string.Empty;
        }

        // Entry point
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a merged range (D4:D5) and set a value
            worksheet.Cells.Merge(3, 3, 2, 1); // firstRow=3, firstColumn=3, totalRows=2, totalColumns=1
            worksheet.Cells[3, 3].PutValue("MergedValue");

            // Prepare a list of custom objects to import
            List<Customer> customers = new List<Customer>
            {
                new Customer { CustomerId = 1, Name = "Customer1" },
                new Customer { CustomerId = 2, Name = "Customer2" },
                new Customer { CustomerId = 3, Name = "Customer3" }
            };

            // Configure import options with CheckMergedCells enabled
            ImportTableOptions options = new ImportTableOptions
            {
                IsFieldNameShown = false,
                InsertRows = true,
                CheckMergedCells = true
            };

            // Import the custom objects starting at cell A1 (row 0, column 0)
            worksheet.Cells.ImportCustomObjects(customers, 0, 0, options);

            // Save the workbook
            workbook.Save("MergedImportDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
