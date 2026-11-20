// Title: Create an Excel workbook, add a named worksheet, and import custom objects while preserving HTML formatting in description cells using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# import custom objects and keep HTML tags in cell values | How to set HtmlString property after ImportCustomObjects in Aspose.Cells | Import POCO list into Excel worksheet with HTML formatting using Aspose.Cells | Preserve bold and italic markup when exporting objects to .xlsx with Aspose.Cells | C# Aspose.Cells example for adding worksheet and writing HTML‑styled descriptions
// Tags: ImportCustomObjects preserve HtmlString Aspose.Cells | HTML markup in Excel cells using Aspose.Cells .NET | C# import POCO objects to worksheet Aspose.Cells | set cell HtmlString property Aspose.Cells | create workbook and named worksheet Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The sample creates a new Workbook, adds a worksheet named "Data", imports an array of custom MyObject instances, assigns each object's HTML‑formatted Description to the corresponding cell via the HtmlString property, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet and give it a name
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet sheet = workbook.Worksheets[sheetIndex];
            sheet.Name = "Data";

            // Sample custom objects with HTML in the Description field
            var items = new List<MyObject>
            {
                new MyObject { Id = 1, Name = "Item 1", Description = "<b>Bold description</b>" },
                new MyObject { Id = 2, Name = "Item 2", Description = "<i>Italic description</i>" },
                new MyObject { Id = 3, Name = "Item 3", Description = "<u>Underlined description</u>" }
            };

            // Import the custom objects starting at row 0, column 0
            // Use default ImportTableOptions (no PreserveFormatting property)
            sheet.Cells.ImportCustomObjects(items.ToArray(), 0, 0, new ImportTableOptions());

            // Apply HTML formatting to the Description column (index 2)
            for (int row = 0; row < items.Count; row++)
            {
                Cell descCell = sheet.Cells[row, 2];
                descCell.HtmlString = items[row].Description; // Preserves HTML formatting
            }

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Definition of the custom object
    public class MyObject
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty; // Contains HTML markup
    }
}
