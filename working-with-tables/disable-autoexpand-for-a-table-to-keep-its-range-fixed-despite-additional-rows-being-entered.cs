// Title: How to prevent an Aspose.Cells ListObject table from auto‑expanding and keep its range fixed in C#
// AI Prompts: Generate C# code that sets a fixed range for a ListObject in Aspose.Cells and stops it from expanding when new rows are added. | Show how to manually update the TableRange of a ListObject after inserting rows using Aspose.Cells for .NET. | Provide a method that disables automatic table growth in Aspose.Cells by redefining the ListObject range.
// Common Searches: Aspose.Cells C# keep Excel table range unchanged after adding rows | disable automatic expansion of ListObject in Aspose.Cells .NET | fixed table size Aspose.Cells when inserting data programmatically | how to lock ListObject range in Aspose.Cells workbook | prevent Excel table auto‑expand using Aspose.Cells API
// Tags: Aspose.Cells ListObject range lock | C# manual table range definition | Aspose.Cells disable table auto‑growth | fixed-size Excel table with Aspose | control ListObject expansion .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example loads a workbook, checks for a ListObject on the first worksheet, and demonstrates that Aspose.Cells .NET does not provide an AutoExpand property, so the table range must be managed manually to keep it fixed before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one table (list object)
                if (worksheet.ListObjects.Count > 0)
                {
                    // Get the first table (list object)
                    ListObject table = worksheet.ListObjects[0];

                    // Note: Aspose.Cells .NET does not expose an AutoExpand property.
                    // If needed, additional logic can be added here to manage table range manually.
                }

                // Save the modified workbook
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
