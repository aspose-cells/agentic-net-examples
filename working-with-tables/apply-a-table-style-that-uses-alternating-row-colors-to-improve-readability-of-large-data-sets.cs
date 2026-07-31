// Title: C# – Apply Alternating Row Colors to an Excel Table with Aspose.Cells (TableStyleMedium2)
// Description: This example creates a new workbook, fills a 100‑row × 5‑column range, converts it to a ListObject (Excel table), applies the built‑in TableStyleMedium2, enables row‑stripe formatting, and saves the result as an XLSX file. It demonstrates how to improve readability of large data sets with alternating row colors using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# alternating row colors | Excel table stripe formatting Aspose | TableStyleMedium2 Aspose.Cells | ShowTableStyleRowStripes example | ListObject style C# | apply built‑in table style .NET | Excel table styling Aspose.Cells
// Common Searches: how to enable row stripes in Aspose.Cells table | C# Aspose.Cells apply TableStyleMedium2 | set ShowTableStyleRowStripes property | alternating row colors for ListObject Aspose | style Excel table programmatically Aspose.Cells
// Developer Intent: Create an Excel table and apply a built‑in style that displays alternating row colors for better visual scanning.
// Use Cases: Generating reports with thousands of rows where stripe formatting aids data review. | Automating export of styled tables to clients while maintaining consistent visual themes. | Batch‑processing multiple worksheets to apply uniform row‑stripe styles across a workbook.
// AI Prompts: Write C# code that builds a ListObject from a range and turns on ShowTableStyleRowStripes using Aspose.Cells. | Explain how to choose different TableStyleType values to get various stripe color schemes in Aspose.Cells. | Show how to customize the stripe colors of a table after applying a built‑in style with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // This example creates a new workbook, fills a 100‑row × 5‑column range, converts it to a ListObject (Excel table), applies the built‑in TableStyleMedium2, enables row‑stripe formatting, and saves the result as an XLSX file. It demonstrates how to improve readability of large data sets with alternating row colors using Aspose.Cells for .NET.
    public class AlternatingRowColorsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate the worksheet with sample data (100 rows, 5 columns)
                for (int col = 0; col < 5; col++)
                {
                    // Header row
                    worksheet.Cells[0, col].PutValue($"Column {col + 1}");
                }

                for (int row = 1; row <= 100; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        worksheet.Cells[row, col].PutValue($"R{row}C{col + 1}");
                    }
                }

                // Convert the range into a table (ListObject)
                // Table range: from A1 to E101 (0‑based indices)
                int tableIndex = worksheet.ListObjects.Add(0, 0, 100, 4, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Apply a built‑in table style that supports row stripes
                table.TableStyleType = TableStyleType.TableStyleMedium2;

                // Enable alternating row colors (row stripe formatting)
                table.ShowTableStyleRowStripes = true;

                // Save the workbook
                string outputPath = "AlternatingRowColors.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AlternatingRowColorsDemo.Run();
        }
    }
}
