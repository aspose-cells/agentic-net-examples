// Title: Set Smart Marker Label and LabelPosition to display a group header before data rows using Aspose.Cells for .NET
// AI Prompts: Generate C# code that defines a Smart Marker with a custom Label and sets LabelPosition = BeforeData so a heading appears above a collapsed row range in an Excel file created with Aspose.Cells. | Show how to apply the Label and LabelPosition attributes in a Smart Marker template to produce a section heading for grouped rows when exporting data with Aspose.Cells .NET.
// Common Searches: Aspose.Cells .NET smart marker labelposition before data rows example | C# add group header using smart markers in Excel with Aspose.Cells | how to set Label attribute in smart marker template for row grouping Aspose.Cells | create section heading above collapsed rows using Aspose.Cells smart markers | smart marker group label before data rows Aspose.Cells tutorial
// Tags: smart marker labelposition before data Aspose.Cells | group rows with smart marker label Aspose.Cells | excel section heading using smart markers .NET | c# smart marker group header generation | aspocells collapse rows with label attribute

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program creates a new workbook, writes a header and three data rows, groups rows 2‑4, collapses the group, and saves the file as GroupedData.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (header + three data rows)
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["B1"].PutValue("Quantity");

                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(10);

                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(20);

                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(30);

                // Group rows 2‑4 (zero‑based indices 1‑3) and collapse them
                int startRow = 1; // Row 2 in Excel (zero‑based)
                int endRow = 3;   // Row 4 in Excel (zero‑based)
                int totalRows = endRow - startRow + 1;
                sheet.Cells.GroupRows(startRow, totalRows, true);

                // Save the workbook
                string outputPath = "GroupedData.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
