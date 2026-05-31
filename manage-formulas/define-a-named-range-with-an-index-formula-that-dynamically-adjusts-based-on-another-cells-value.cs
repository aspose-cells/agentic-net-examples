using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

namespace AsposeCellsDynamicNamedRange
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (Sheet1)
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";
                Cells cells = sheet.Cells;

                // Populate range A1:A10 with sample data (1 to 10)
                for (int i = 0; i < 10; i++)
                {
                    cells[i, 0].PutValue(i + 1); // Column A
                }

                // Cell B1 will hold the index that determines which item to pick
                cells["B1"].PutValue(3); // Initial index = 3 (will pick value 3)

                // Add a named range that uses INDEX to return a single cell from A1:A10
                int nameIndex = workbook.Worksheets.Names.Add("SelectedItem");
                Name dynamicName = workbook.Worksheets.Names[nameIndex];
                // The formula refers to the index value in B1; it will adjust automatically
                dynamicName.RefersTo = "=INDEX(Sheet1!$A$1:$A$10, Sheet1!$B$1)";

                // Use the named range in a formula (C1 will display the selected value)
                cells["C1"].Formula = "=SelectedItem";

                // Calculate formulas
                workbook.CalculateFormula();

                // Output the result after first calculation
                Console.WriteLine("Initial index (B1=3) => SelectedItem value: " + cells["C1"].Value);

                // Change the index to 7 and recalculate
                cells["B1"].PutValue(7);
                workbook.CalculateFormula();

                // Output the updated result
                Console.WriteLine("Updated index (B1=7) => SelectedItem value: " + cells["C1"].Value);

                // Demonstrate retrieving the range that the name resolves to
                // Since the name refers to a single cell, GetRange will return that cell's range
                AsposeRange resolvedRange = dynamicName.GetRange();
                Console.WriteLine("Resolved range address: " + resolvedRange.RefersTo);

                // Save the workbook (ensure the directory exists)
                string outputPath = "DynamicNamedRangeDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}