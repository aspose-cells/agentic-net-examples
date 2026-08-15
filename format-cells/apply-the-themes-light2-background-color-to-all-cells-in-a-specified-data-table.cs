// Title: Apply Light2 Theme Background to an Excel Table with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, fills cells A1:C4, converts the range into a ListObject, and applies the built‑in TableStyleLight2 so the entire table inherits the workbook's Light2 theme background color. The workbook is then saved as an .xlsx file.
// Keywords: Aspose.Cells Light2 table style | C# apply Excel theme background | TableStyleLight2 Aspose.Cells example | set built‑in table style .NET | Excel ListObject theme color Aspose | apply theme background to table programmatically | Aspose.Cells formatting tables | C# Excel table style Light2
// Common Searches: Aspose.Cells apply Light2 style to table | C# set TableStyleLight2 for ListObject | how to use theme colors in Aspose.Cells | apply built‑in Excel table style with Aspose.Cells .NET | change Excel table background to Light2 using code
// Developer Intent: Apply the built‑in Light2 table style so the whole ListObject uses the workbook’s Light2 theme background color.
// Use Cases: Generate sales or inventory reports that automatically match the workbook’s Light2 theme for consistent branding. | Export processed data to Excel with a ready‑made Light2 styled table to improve readability. | Create templates where all tables adopt the Light2 background, reducing manual formatting effort.
// AI Prompts: Show how to switch the table style from Light2 to Light3 after the ListObject is created with Aspose.Cells. | Provide code to set a custom column background while preserving the Light2 table style for the rest of the table. | Explain how to read the actual RGB value of the workbook’s Light2 theme color using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // This example creates a workbook, fills cells A1:C4, converts the range into a ListObject, and applies the built‑in TableStyleLight2 so the entire table inherits the workbook's Light2 theme background color. The workbook is then saved as an .xlsx file.
    public class ApplyLight2BackgroundToTable
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data for the table (A1:C4)
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Category");
                cells["C1"].PutValue("Price");

                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue("Fruit");
                cells["C2"].PutValue(1.2);

                cells["A3"].PutValue("Carrot");
                cells["B3"].PutValue("Vegetable");
                cells["C3"].PutValue(0.8);

                cells["A4"].PutValue("Bread");
                cells["B4"].PutValue("Grain");
                cells["C4"].PutValue(2.5);

                // Create a ListObject (Excel table) covering the data range
                int tableIndex = worksheet.ListObjects.Add("A1", "C4", true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Apply the built‑in Light2 table style which uses the theme's Light2 background color
                table.TableStyleType = TableStyleType.TableStyleLight2;

                // Ensure the style is applied to the entire range of the table
                table.ApplyStyleToRange();

                // Save the workbook
                string outputPath = "ApplyLight2BackgroundToTable.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyLight2BackgroundToTable.Run();
        }
    }
}
