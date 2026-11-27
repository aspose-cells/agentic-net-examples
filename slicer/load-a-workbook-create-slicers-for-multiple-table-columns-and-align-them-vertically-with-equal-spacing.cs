// Title: C# example: add multiple slicers to an Excel table and align them vertically with equal spacing using Aspose.Cells
// Common Searches: Aspose.Cells C# add slicers to a ListObject and set their positions | How to align Excel slicers vertically with the same size using Aspose.Cells .NET | Programmatically create slicers for several table columns in C# with Aspose.Cells | Set slicer width, height and row gap in an Excel workbook via Aspose.Cells | C# code to add slicers to a table and arrange them with equal spacing
// Tags: Aspose.Cells add slicer to ListObject C# | Aspose.Cells set slicer size Excel | Aspose.Cells vertical slicer alignment | Aspose.Cells slicer row spacing | Aspose.Cells create table slicers programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

namespace SlicerAlignmentDemo
{
    // The sample loads an existing workbook (or creates one), ensures a ListObject table exists, adds two slicers for the first two columns, positions them with a defined row offset, applies a uniform width and height, gives each slicer a descriptive name, and saves the workbook with the vertically aligned slicers.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output_with_slicers.xlsx";

            // Ensure Aspose.Cells license is set if you have one.
            // License license = new License();
            // license.SetLicense("Aspose.Cells.lic");

            Workbook workbook = null;

            try
            {
                // Load existing workbook or create a new one with sample data and a table
                if (File.Exists(inputPath))
                {
                    Console.WriteLine($"Loading workbook from \"{inputPath}\"...");
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file not found. Creating a sample workbook \"{inputPath}\"...");
                    workbook = CreateSampleWorkbook();
                    workbook.Save(inputPath);
                    Console.WriteLine($"Sample workbook saved to \"{inputPath}\".");
                }

                // Work with the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure there is at least one table (ListObject) on the sheet
                if (sheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("No tables found on the worksheet. Creating a sample table...");
                    int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
                    ListObject table = sheet.ListObjects[tableIndex];

                    // Populate sample data for the table
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["B1"].PutValue("Region");
                    sheet.Cells["C1"].PutValue("Sales");
                    sheet.Cells["A2"].PutValue("Electronics");
                    sheet.Cells["B2"].PutValue("North");
                    sheet.Cells["C2"].PutValue(1200);
                    sheet.Cells["A3"].PutValue("Electronics");
                    sheet.Cells["B3"].PutValue("South");
                    sheet.Cells["C3"].PutValue(850);
                    sheet.Cells["A4"].PutValue("Furniture");
                    sheet.Cells["B4"].PutValue("North");
                    sheet.Cells["C4"].PutValue(430);
                    sheet.Cells["A5"].PutValue("Furniture");
                    sheet.Cells["B5"].PutValue("South");
                    sheet.Cells["C5"].PutValue(670);
                    Console.WriteLine("Sample table created.");
                }

                // Retrieve the first table (ListObject)
                ListObject listTable = sheet.ListObjects[0];

                // Verify that the table has at least two columns for slicers
                if (listTable.ListColumns.Count < 2)
                {
                    Console.WriteLine("The table does not have enough columns to create multiple slicers.");
                    return;
                }

                ListColumn column1 = listTable.ListColumns[0]; // e.g., "Category"
                ListColumn column2 = listTable.ListColumns[1]; // e.g., "Region"

                // Define slicer layout parameters
                const double slicerWidth = 150.0;   // points
                const double slicerHeight = 200.0;  // points
                const int startRow = 1;             // zero‑based row index
                const int startColumn = 1;          // zero‑based column index
                const int rowSpacing = 12;          // rows between slicers

                // Add slicers at the desired positions
                int slicerIdx1 = sheet.Slicers.Add(listTable, column1, startRow, startColumn);
                int slicerIdx2 = sheet.Slicers.Add(listTable, column2, startRow + rowSpacing, startColumn);

                Slicer slicer1 = sheet.Slicers[slicerIdx1];
                Slicer slicer2 = sheet.Slicers[slicerIdx2];

                // Set uniform size using Width and Height properties
                slicer1.Width = slicerWidth;
                slicer1.Height = slicerHeight;
                slicer2.Width = slicerWidth;
                slicer2.Height = slicerHeight;

                // Optional: give slicers meaningful names
                slicer1.Name = "Slicer_" + column1.Name;
                slicer2.Name = "Slicer_" + column2.Name;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with slicers to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static Workbook CreateSampleWorkbook()
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("Product");
            ws.Cells["B1"].PutValue("Category");
            ws.Cells["C1"].PutValue("Units Sold");
            ws.Cells["A2"].PutValue("Laptop");
            ws.Cells["B2"].PutValue("Electronics");
            ws.Cells["C2"].PutValue(150);
            ws.Cells["A3"].PutValue("Desk");
            ws.Cells["B3"].PutValue("Furniture");
            ws.Cells["C3"].PutValue(85);
            ws.Cells["A4"].PutValue("Smartphone");
            ws.Cells["B4"].PutValue("Electronics");
            ws.Cells["C4"].PutValue(230);
            ws.Cells["A5"].PutValue("Chair");
            ws.Cells["B5"].PutValue("Furniture");
            ws.Cells["C5"].PutValue(120);

            // Convert the range into a table (ListObject)
            ws.ListObjects.Add(0, 0, 4, 2, true);
            return wb;
        }
    }
}
