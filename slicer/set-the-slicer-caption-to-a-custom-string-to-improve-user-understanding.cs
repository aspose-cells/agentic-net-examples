// Title: Aspose.Cells C# – Set a Custom Caption for a Pivot Table Slicer
// Description: Learn how to create a workbook, add sample data, build a pivot table, insert a slicer for the "Category" field, and assign a custom caption (e.g., "Select Category") using Aspose.Cells for .NET. The example ensures the caption header is visible and saves the file as an Excel workbook.
// Keywords: Aspose.Cells slicer caption | C# set slicer header | custom slicer caption .NET | pivot table slicer Aspose.Cells | Aspose.Cells change slicer title | Excel slicer caption programmatically | Aspose.Cells C# example
// Common Searches: how to change slicer caption Aspose.Cells C# | set custom caption for pivot slicer using Aspose.Cells | Aspose.Cells slicer caption example | C# code to add slicer and set caption in Excel | Aspose.Cells display slicer header
// Developer Intent: Add a slicer to a pivot table and define a custom caption for it with Aspose.Cells for .NET.
// Use Cases: Provide a clear, user‑friendly label for a slicer that filters a specific field in generated reports. | Create Excel dashboards where slicer headers describe their purpose instead of showing the raw field name. | Automate workbook generation with pivot‑based analysis and customized slicer captions for better end‑user experience.
// AI Prompts: Show C# code using Aspose.Cells to add a slicer to a pivot table and set a custom caption. | Explain how to make the slicer caption visible and saved in an Excel file with Aspose.Cells. | Provide step‑by‑step instructions to create a workbook, build a pivot table, attach a slicer, and customize its caption in .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable and PivotFieldType
using Aspose.Cells.Slicers; // Required for Slicer

namespace AsposeCellsExample
{
    // Learn how to create a workbook, add sample data, build a pivot table, insert a slicer for the "Category" field, and assign a custom caption (e.g., "Select Category") using Aspose.Cells for .NET. The example ensures the caption header is visible and saves the file as an Excel workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate worksheet with sample data for a pivot table
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Fruit");
                worksheet.Cells["A3"].PutValue("Vegetable");
                worksheet.Cells["A4"].PutValue("Fruit");
                worksheet.Cells["A5"].PutValue("Vegetable");

                worksheet.Cells["B1"].PutValue("Amount");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(80);
                worksheet.Cells["B4"].PutValue(150);
                worksheet.Cells["B5"].PutValue(90);

                // Add a pivot table based on the data range A1:B5, place it at D1
                int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Add a slicer linked to the pivot table for the "Category" field, place it at E1
                int slicerIndex = worksheet.Slicers.Add(pivotTable, "Category", "E1");
                Slicer slicer = worksheet.Slicers[slicerIndex];

                // Set a custom caption for the slicer
                slicer.Caption = "Select Category";
                slicer.ShowCaption = true; // Ensure the caption header is visible

                // Define output file path
                string outputPath = "SlicerWithCustomCaption.xlsx";

                // Ensure the directory exists (handle case where outputPath has no directory)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
