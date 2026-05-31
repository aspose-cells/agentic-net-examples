using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangesDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Add sample named ranges for demonstration
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells.CreateRange("A1:B2").Name = "SalesData";
                sheet.Cells.CreateRange("C3:D4").Name = "Expenses";

                // Retrieve all defined named ranges (feature rule: GetNamedRanges)
                // Use fully qualified type to avoid conflict with System.Range
                Aspose.Cells.Range[] namedRanges = workbook.Worksheets.GetNamedRanges();

                // Output each named range's name and address to the console
                if (namedRanges != null && namedRanges.Length > 0)
                {
                    Console.WriteLine($"Found {namedRanges.Length} named ranges:");
                    foreach (Aspose.Cells.Range range in namedRanges)
                    {
                        Console.WriteLine($"Name: {range.Name}, Address: {range.Address}");
                    }
                }
                else
                {
                    Console.WriteLine("No named ranges found.");
                }

                // Save the workbook (lifecycle rule: save)
                string outputPath = "NamedRangesDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}