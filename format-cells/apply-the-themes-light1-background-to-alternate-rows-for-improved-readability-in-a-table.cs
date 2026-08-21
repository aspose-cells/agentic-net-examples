// Title: Apply Light1 Table Style with Alternating Row Stripes in Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, insert a ListObject, set the built‑in TableStyleLight1, enable row‑stripe formatting, and save the file as XLSX using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | TableStyleLight1 | row stripes | alternating row background | ListObject | Excel table formatting | SaveFormat.Xlsx | built‑in table style
// Common Searches: Aspose.Cells TableStyleLight1 C# example | Enable row stripe formatting Aspose.Cells | How to add ListObject with style in .NET | Apply built‑in table style with alternate rows | C# Aspose.Cells alternating row colors
// Developer Intent: Apply the Light1 built‑in table style and turn on row‑stripe formatting for a worksheet table using Aspose.Cells.
// Use Cases: Create readable data tables in reports with header shading and alternating row colors. | Generate Excel exports that follow a corporate Light1 theme while improving scanability. | Produce large‑scale spreadsheets where visual row separation aids data analysis.
// AI Prompts: Generate C# code that adds a ListObject, sets TableStyleLight1, enables ShowTableStyleRowStripes, and saves the workbook with Aspose.Cells. | Explain the impact of ShowTableStyleRowStripes on Excel tables and how it works with built‑in styles in Aspose.Cells. | Provide a step‑by‑step guide to apply Light1 style with alternating row backgrounds in a .NET Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, insert a ListObject, set the built‑in TableStyleLight1, enable row‑stripe formatting, and save the file as XLSX using Aspose.Cells for .NET.
    public class ApplyLight1RowStripes
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the table (A1:B5)
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("Item1");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["A3"].PutValue("Item2");
            worksheet.Cells["B3"].PutValue(200);
            worksheet.Cells["A4"].PutValue("Item3");
            worksheet.Cells["B4"].PutValue(300);
            worksheet.Cells["A5"].PutValue("Item4");
            worksheet.Cells["B5"].PutValue(400);

            // Add a ListObject (table) covering the data range
            int tableIndex = worksheet.ListObjects.Add("A1", "B5", true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Apply the built‑in Light1 table style
            table.TableStyleType = TableStyleType.TableStyleLight1;

            // Enable row stripe formatting (alternating background)
            table.ShowTableStyleRowStripes = true;

            // Save the workbook
            workbook.Save("TableWithLight1RowStripes.xlsx", SaveFormat.Xlsx);
        }
    }
}
