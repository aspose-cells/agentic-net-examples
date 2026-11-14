// Title: Add category‑specific images to Excel cells with Aspose.Cells C# using conditional If logic
// Common Searches: Aspose.Cells C# insert image into cell based on column value | how to display different pictures for product categories in Excel using Aspose.Cells | C# conditional picture insertion in worksheet with Aspose.Cells API | add PNG image to Excel row when category equals Electronics using Aspose.Cells | Aspose.Cells example for mapping product category to picture in Excel
// Tags: conditional image insertion Aspose.Cells C# | add picture to worksheet cell Aspose.Cells | map product category to PNG in Excel using Aspose.Cells | Aspose.Cells picture handling C# example | Excel cell image based on value Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, writes product categories in column A, then iterates rows and inserts a matching PNG image into column B using Aspose.Cells' picture API and If‑based conditional logic, finally saving the file as an Excel workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data: Product Category in column A (starting from row 1)
                sheet.Cells["A1"].PutValue("Product Category");
                sheet.Cells["A2"].PutValue("Electronics");
                sheet.Cells["A3"].PutValue("Clothing");
                sheet.Cells["A4"].PutValue("Food");
                sheet.Cells["A5"].PutValue("Electronics");

                // Paths to picture files
                string electronicsPath = "Electronics.png";
                string clothingPath = "Clothing.png";
                string foodPath = "Food.png";

                // Helper to add a picture at a specific cell if the file exists
                void AddPictureAt(int row, int column, string path)
                {
                    try
                    {
                        if (!File.Exists(path))
                        {
                            Console.WriteLine($"Picture file not found: {path}");
                            return;
                        }

                        int index = sheet.Pictures.Add(row, column, path);
                        Picture picture = sheet.Pictures[index];
                        picture.Width = 80;
                        picture.Height = 80;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to add picture '{path}' at ({row},{column}): {ex.Message}");
                    }
                }

                // Apply conditional logic: add the picture that matches the product category
                int firstDataRow = 1; // Excel row 2 (zero‑based)
                int lastDataRow = sheet.Cells.MaxDataRow; // zero‑based index of last row with data

                for (int row = firstDataRow; row <= lastDataRow; row++)
                {
                    string category = sheet.Cells[row, 0].StringValue; // Column A

                    if (category.Equals("Electronics", StringComparison.OrdinalIgnoreCase))
                    {
                        AddPictureAt(row, 1, electronicsPath); // Column B
                    }
                    else if (category.Equals("Clothing", StringComparison.OrdinalIgnoreCase))
                    {
                        AddPictureAt(row, 1, clothingPath);
                    }
                    else if (category.Equals("Food", StringComparison.OrdinalIgnoreCase))
                    {
                        AddPictureAt(row, 1, foodPath);
                    }
                }

                // Save the workbook
                string outputPath = "ProductCategoriesWithImages.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
