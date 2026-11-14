// Title: Add a totals row to an Aspose.Cells ListObject and configure Count aggregation for a text column in C#
// Common Searches: how to add a totals row to a ListObject using Aspose.Cells C# | Aspose.Cells C# count aggregation for a text column in an Excel table | set totals calculation for a specific column in an Aspose.Cells worksheet | enable totals row and count totals for a category column with Aspose.Cells | C# example of configuring ListColumns TotalsCalculation property
// Tags: Aspose.Cells ListObject totals row | C# configure column totals calculation | Aspose.Cells count aggregation for text column | Excel table totals row using Aspose.Cells | ListColumns TotalsCalculation property C# | create ListObject with totals in Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example creates a workbook, defines a range as a ListObject (Excel table) with headers, enables the totals row, sets the TotalsCalculation of the text "Category" column to Count, and saves the file as ListObjectWithTotals.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (including a text column)
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Category"); // Text column we will aggregate
                sheet.Cells["C1"].PutValue("Amount");

                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alpha");
                sheet.Cells["C2"].PutValue(100);

                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Beta");
                sheet.Cells["C3"].PutValue(150);

                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue("Alpha");
                sheet.Cells["C4"].PutValue(200);

                // Define the range that will become a ListObject (table)
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
                int firstRow = 0;          // Row index for header (0‑based)
                int firstColumn = 0;       // Column index for header (0‑based)
                int totalRows = 4;         // Total rows including header
                int totalColumns = 3;      // Number of columns
                bool hasHeaders = true;

                // Add the ListObject to the worksheet
                int listObjectIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, hasHeaders);
                ListObject listObject = sheet.ListObjects[listObjectIndex];

                // Enable the Totals row for the ListObject
                listObject.ShowTotals = true;

                // Configure the "Category" (text) column to use Count aggregation in the Totals row
                // Assuming "Category" is the second column (index 1)
                listObject.ListColumns[1].TotalsCalculation = TotalsCalculation.Count;

                // Define output file path
                string outputPath = "ListObjectWithTotals.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{System.IO.Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
