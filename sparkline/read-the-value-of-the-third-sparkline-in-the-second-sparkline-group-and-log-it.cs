// Title: Read the DataRange of the third sparkline in the second sparkline group with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds two line sparkline groups (each containing three sparklines), accesses the third sparkline of the second group, writes its DataRange to the console, and saves the file. It demonstrates how to retrieve a specific sparkline's source range using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# sparkline example | sparkline DataRange | second sparkline group | third sparkline | read sparkline range | log sparkline data | Aspose.Cells .NET tutorial | Excel sparkline API | access sparkline programmatically
// Common Searches: Aspose.Cells get sparkline DataRange C# | How to read third sparkline in second group | Access specific sparkline with Aspose.Cells | C# example for sparkline group indexing | Retrieve sparkline source range Aspose
// Developer Intent: Obtain and display the DataRange of a particular sparkline within a sparkline group.
// Use Cases: Debugging sparkline source ranges before publishing a workbook | Validating that sparklines reference the intended cells | Collecting sparkline metadata for reporting or analytics | Automating quality checks on generated Excel files
// AI Prompts: Generate C# code that loops through all sparkline groups in a worksheet and prints each sparkline's DataRange. | Show how to change the DataRange of the third sparkline in the second sparkline group to a new cell range using Aspose.Cells. | Provide an example that extracts numeric values from a sparkline's DataRange and calculates their total in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineDemo
{
    // This example creates a workbook, adds two line sparkline groups (each containing three sparklines), accesses the third sparkline of the second group, writes its DataRange to the console, and saves the file. It demonstrates how to retrieve a specific sparkline's source range using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate data for two rows (each row will be a sparkline group)
                worksheet.Cells["A1"].PutValue(1);
                worksheet.Cells["B1"].PutValue(2);
                worksheet.Cells["C1"].PutValue(3);
                worksheet.Cells["D1"].PutValue(4);

                worksheet.Cells["A2"].PutValue(5);
                worksheet.Cells["B2"].PutValue(6);
                worksheet.Cells["C2"].PutValue(7);
                worksheet.Cells["D2"].PutValue(8);

                // ---------- First Sparkline Group ----------
                // Location of the first sparkline (single cell E1)
                CellArea location1 = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 4,
                    EndColumn = 4
                };
                int groupIndex1 = worksheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location1);
                SparklineGroup group1 = worksheet.SparklineGroups[groupIndex1];

                // Add three sparklines to the first group (different columns)
                group1.Sparklines.Add(worksheet.Name + "!A1:D1", 0, 4);
                group1.Sparklines.Add(worksheet.Name + "!A1:D1", 0, 5);
                group1.Sparklines.Add(worksheet.Name + "!A1:D1", 0, 6);

                // ---------- Second Sparkline Group ----------
                // Location of the first sparkline in the second group (single cell E2)
                CellArea location2 = new CellArea
                {
                    StartRow = 1,
                    EndRow = 1,
                    StartColumn = 4,
                    EndColumn = 4
                };
                int groupIndex2 = worksheet.SparklineGroups.Add(SparklineType.Line, "A2:D2", false, location2);
                SparklineGroup group2 = worksheet.SparklineGroups[groupIndex2];

                // Add three sparklines to the second group
                group2.Sparklines.Add(worksheet.Name + "!A2:D2", 1, 4);
                group2.Sparklines.Add(worksheet.Name + "!A2:D2", 1, 5);
                group2.Sparklines.Add(worksheet.Name + "!A2:D2", 1, 6);

                // Access the third sparkline (index 2) in the second sparkline group (index 1)
                Sparkline thirdSparklineInSecondGroup = worksheet.SparklineGroups[1].Sparklines[2];

                // Log its DataRange (the range of cells the sparkline visualises)
                Console.WriteLine("Third sparkline in second group DataRange: " + thirdSparklineInSecondGroup.DataRange);

                // Save the workbook
                workbook.Save("SparklineDemo.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
