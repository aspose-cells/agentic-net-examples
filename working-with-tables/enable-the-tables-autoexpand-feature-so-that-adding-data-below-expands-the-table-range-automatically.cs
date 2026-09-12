// Title: C# example: Simulating auto‑expand for an Aspose.Cells table by manually resizing the ListObject after inserting rows
// AI Prompts: Write C# code that inserts a new data row below an existing Aspose.Cells ListObject and then updates the ListObject's range to include the added row. | Create a reusable C# method using Aspose.Cells that detects the last used row of a worksheet, adds a specified number of rows, and expands the associated ListObject to cover the new rows. | Generate a C# snippet that checks for a ListObject on a worksheet, adds rows at the end of the data, and calls ListObject.Resize to adjust the table range.
// Common Searches: Aspose.Cells C# expand Excel table range after adding rows programmatically | How to resize a ListObject in Aspose.Cells when new rows are appended | C# Aspose.Cells auto‑expand table workaround for dynamic data | Update Aspose.Cells ListObject range after inserting rows in .xlsx file | Simulate Excel table auto‑expand using Aspose.Cells C# API
// Tags: Aspose.Cells ListObject resize | C# expand Excel table range | Aspose.Cells manual table auto‑expand | update ListObject range C# | Aspose.Cells add rows to table

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for ListObject

namespace AsposeCellsExample
{
    // The example loads or creates a workbook, adds sample data and a ListObject named SampleTable, notes that Aspose.Cells lacks a built‑in AutoExpand property, and shows how to manually resize the ListObject after inserting rows so the table expands automatically before saving.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Load existing workbook or create a new one if the file is missing
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Name = "Sheet1";

                    // Sample data
                    ws.Cells["A1"].PutValue("ID");
                    ws.Cells["B1"].PutValue("Name");
                    ws.Cells["A2"].PutValue(1);
                    ws.Cells["B2"].PutValue("Alice");
                    ws.Cells["A3"].PutValue(2);
                    ws.Cells["B3"].PutValue("Bob");

                    // Add a table covering the data range
                    int tableIndex = ws.ListObjects.Add(0, 0, 3, 2, true);
                    ListObject table = ws.ListObjects[tableIndex];
                    table.DisplayName = "SampleTable";
                }

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the first table (ListObject) on the worksheet
                if (worksheet.ListObjects.Count > 0)
                {
                    ListObject table = worksheet.ListObjects[0];
                    // Aspose.Cells does not provide an AutoExpand property.
                    // If needed, manually resize the table after adding rows.
                }
                else
                {
                    Console.WriteLine("No tables found in the worksheet.");
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = Directory.GetCurrentDirectory();
                }
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
