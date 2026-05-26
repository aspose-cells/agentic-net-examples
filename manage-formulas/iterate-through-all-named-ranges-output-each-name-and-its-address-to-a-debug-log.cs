using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsNamedRangesDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add some sample data
                sheet.Cells["A1"].PutValue("Item1");
                sheet.Cells["A2"].PutValue("Item2");
                sheet.Cells["B1"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);

                // Create named ranges for demonstration
                sheet.Cells.CreateRange("A1:A2").Name = "ItemNames";
                sheet.Cells.CreateRange("B1:B2").Name = "ItemValues";

                // Retrieve all named ranges in the workbook
                Aspose.Cells.Range[] namedRanges = workbook.Worksheets.GetNamedRanges();

                // If there are no named ranges, log a message
                if (namedRanges == null || namedRanges.Length == 0)
                {
                    Debug.WriteLine("No named ranges found in the workbook.");
                    return;
                }

                // Iterate through each named range and output its name and address
                foreach (Aspose.Cells.Range range in namedRanges)
                {
                    // range.Name holds the defined name, range.Address holds the address (e.g., A1:A2)
                    Debug.WriteLine($"Name: {range.Name}, Address: {range.Address}");
                }

                // Optionally save the workbook (demonstrates lifecycle usage)
                string outputPath = "NamedRangesDemo.xlsx";
                workbook.Save(outputPath);
                Debug.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Debug.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}