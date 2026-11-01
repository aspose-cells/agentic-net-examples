// Title: How to insert category‑specific images into Excel rows with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add picture based on cell content | C# loop through Excel rows and display different PNGs with Aspose.Cells | How to use a dictionary to select images for product categories in Aspose.Cells | Conditional image rendering in .xlsx files using Aspose.Cells for .NET | Example of inserting category specific pictures into Excel using Aspose.Cells C#
// Tags: conditional picture insertion Aspose.Cells | category image mapping C# | insert picture into worksheet cell Aspose.Cells | excel row image rendering .xlsx | dictionary driven image lookup Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ConditionalImageDemo
{
    // The sample creates a workbook, adds product and category columns, defines a case‑insensitive dictionary that maps category names to PNG files, iterates over each data row, and inserts the matching picture into column C using the Aspose.Cells Pictures.Add method. Missing images are logged, and the final file is saved as ProductsWithConditionalImages.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add header row.
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Category");

                // Sample data.
                sheet.Cells["A2"].PutValue("Laptop");
                sheet.Cells["B2"].PutValue("Electronics");
                sheet.Cells["A3"].PutValue("T‑Shirt");
                sheet.Cells["B3"].PutValue("Clothing");
                sheet.Cells["A4"].PutValue("Apple");
                sheet.Cells["B4"].PutValue("Food");

                // Define the mapping between category and image file.
                var categoryImages = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "Electronics", "Electronics.png" },
                    { "Clothing",    "Clothing.png" },
                    { "Food",        "Food.png" }
                };

                // Insert the appropriate picture for each data row.
                int firstDataRow = 1; // zero‑based index (row 2 in Excel)
                int lastDataRow = sheet.Cells.MaxDataRow;

                for (int row = firstDataRow; row <= lastDataRow; row++)
                {
                    string category = sheet.Cells[row, 1].StringValue; // Column B (index 1)

                    if (categoryImages.TryGetValue(category, out string imagePath) && File.Exists(imagePath))
                    {
                        // Add picture anchored to the cell in column C (index 2) of the same row.
                        // Adjust the picture's placement as needed.
                        sheet.Pictures.Add(row, 2, imagePath);
                    }
                    else
                    {
                        // Image file missing – optionally log or ignore.
                        Console.WriteLine($"Image for category \"{category}\" not found at path \"{imagePath}\".");
                    }
                }

                // Save the workbook.
                string outputPath = "ProductsWithConditionalImages.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
