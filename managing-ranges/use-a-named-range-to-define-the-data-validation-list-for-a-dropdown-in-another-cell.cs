using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class NamedRangeDropDownDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate the source list in column A (A1:A4)
                worksheet.Cells["A1"].PutValue("Apple");
                worksheet.Cells["A2"].PutValue("Banana");
                worksheet.Cells["A3"].PutValue("Cherry");
                worksheet.Cells["A4"].PutValue("Date");

                // Create a named range called "FruitList" that refers to A1:A4
                int nameIndex = workbook.Worksheets.Names.Add("FruitList");
                workbook.Worksheets.Names[nameIndex].RefersTo = $"={worksheet.Name}!$A$1:$A$4";

                // Add data validation to cell B1 and use the named range as the list source
                Validation validation = worksheet.Cells["B1"].GetValidation();
                validation.Type = ValidationType.List;      // List type validation
                validation.Formula1 = "FruitList";          // Reference the named range (no leading '=')
                validation.InCellDropDown = true;           // Show the drop‑down arrow in the cell

                // Save the workbook
                string outputPath = "NamedRangeDropDown.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            NamedRangeDropDownDemo.Run();
        }
    }
}