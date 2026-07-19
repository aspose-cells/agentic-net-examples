// Title: C# – Filter a ListObject (Excel Table) where Region = "East" using Aspose.Cells
// Description: This Aspose.Cells for .NET example creates a workbook, adds a ListObject with ID, Region and Sales columns, enables AutoFilter, finds the Region column index, applies a filter to show only rows with Region set to "East", refreshes the view and saves the file.
// Keywords: Aspose.Cells C# filter ListObject | AutoFilter Excel table .NET | filter rows by column value Aspose.Cells | ListObject region filter example | C# Excel table AutoFilter | Aspose.Cells table column index | region east filter Aspose.Cells | Excel workbook filtering C#
// Common Searches: Aspose.Cells filter ListObject by column value | C# apply AutoFilter to Excel table | How to filter rows where Region = East in Aspose.Cells | Get column index of header in ListObject Aspose.Cells | Refresh AutoFilter after applying filter C#
// Developer Intent: Apply an AutoFilter to a ListObject so that only rows with the Region column equal to "East" remain visible.
// Use Cases: Produce a sales report that automatically displays only Eastern‑region records before distribution. | Build a dynamic dashboard that toggles data views by filtering the Region column of an Excel table. | Export a subset of a large worksheet by filtering for a specific region and saving the filtered workbook.
// AI Prompts: Generate C# code with Aspose.Cells that filters a ListObject to show rows where a given column equals a specific string. | Explain how to locate a header’s relative index inside a ListObject and use it to apply an AutoFilter. | Show an Aspose.Cells example that applies multiple criteria (e.g., Region = "East" and Sales > 1000) to a ListObject.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // This Aspose.Cells for .NET example creates a workbook, adds a ListObject with ID, Region and Sales columns, enables AutoFilter, finds the Region column index, applies a filter to show only rows with Region set to "East", refreshes the view and saves the file.
    public class ListObjectRegionFilterDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data with a "Region" column
                worksheet.Cells["A1"].PutValue("ID");
                worksheet.Cells["B1"].PutValue("Region");
                worksheet.Cells["C1"].PutValue("Sales");

                worksheet.Cells["A2"].PutValue(1);
                worksheet.Cells["B2"].PutValue("East");
                worksheet.Cells["C2"].PutValue(1200);

                worksheet.Cells["A3"].PutValue(2);
                worksheet.Cells["B3"].PutValue("West");
                worksheet.Cells["C3"].PutValue(850);

                worksheet.Cells["A4"].PutValue(3);
                worksheet.Cells["B4"].PutValue("East");
                worksheet.Cells["C4"].PutValue(950);

                worksheet.Cells["A5"].PutValue(4);
                worksheet.Cells["B5"].PutValue("South");
                worksheet.Cells["C5"].PutValue(400);

                // Add a ListObject (table) that covers the data range (including header)
                int listObjectIndex = worksheet.ListObjects.Add(0, 0, 4, 2, true);
                ListObject listObject = worksheet.ListObjects[listObjectIndex];

                // Ensure the table has an AutoFilter
                listObject.HasAutoFilter = true;

                // Determine the column index of the "Region" header (relative to the table)
                int regionColumnIndex = -1;
                for (int col = listObject.StartColumn; col <= listObject.EndColumn; col++)
                {
                    string header = worksheet.Cells[listObject.StartRow, col].StringValue;
                    if (header.Equals("Region", StringComparison.OrdinalIgnoreCase))
                    {
                        regionColumnIndex = col - listObject.StartColumn; // relative index
                        break;
                    }
                }

                if (regionColumnIndex == -1)
                    throw new Exception("Region column not found.");

                // Apply filter to show only rows where Region = "East"
                listObject.AutoFilter.Filter(regionColumnIndex, "East");
                listObject.AutoFilter.Refresh();

                // Save the workbook
                string outputPath = "ListObjectRegionFilterDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ListObjectRegionFilterDemo.Run();
        }
    }
}
