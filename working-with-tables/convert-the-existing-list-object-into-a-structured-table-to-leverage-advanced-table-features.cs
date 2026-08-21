// Title: Convert a Cell Range to a Styled ListObject Table with Totals Using Aspose.Cells for .NET (C#)
// Description: This example creates a new workbook, fills cells A1:C5 with product data, converts the range into a ListObject named "ProductTable", applies the TableStyleMedium9 style, shows a totals row, sets the Price column to calculate a sum, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# convert range to table | Aspose.Cells ListObject | Excel table style Aspose | Add totals row Aspose.Cells | Sum calculation column Aspose | Create structured table .NET | Apply TableStyleMedium9 | Workbook.Save Excel
// Common Searches: how to convert a worksheet range to a ListObject table using Aspose.Cells | Aspose.Cells .NET add totals row and sum calculation | apply predefined table style to a range with Aspose.Cells | set display name for ListObject in Aspose.Cells C# | convert range to table Aspose.Cells example
// Developer Intent: Transform a plain cell range into a named ListObject table, apply a built‑in style, and enable a totals row with sum calculation in a .NET workbook.
// Use Cases: Generate a product catalog where raw data is automatically formatted as a styled table with a total price row. | Prepare financial statements by converting a data range into a ListObject and adding sum totals for amount columns. | Export data for downstream systems that require a named Excel table for formula references or Power Query ingestion.
// AI Prompts: Show how to add a custom label to the totals row in the ListObject. | Provide code that changes the totals calculation of a Quantity column to Average and formats the total row as currency. | Explain how to retrieve a ListObject by its DisplayName and modify its style after the workbook has been saved.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // This example creates a new workbook, fills cells A1:C5 with product data, converts the range into a ListObject named "ProductTable", applies the TableStyleMedium9 style, shows a totals row, sets the Price column to calculate a sum, and saves the file as an Excel workbook.
    public class ConvertRangeToTableDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data in a plain range (A1:C5)
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Category");
                sheet.Cells["C1"].PutValue("Price");

                sheet.Cells["A2"].PutValue("Laptop");
                sheet.Cells["B2"].PutValue("Electronics");
                sheet.Cells["C2"].PutValue(1200);

                sheet.Cells["A3"].PutValue("Phone");
                sheet.Cells["B3"].PutValue("Electronics");
                sheet.Cells["C3"].PutValue(800);

                sheet.Cells["A4"].PutValue("Desk");
                sheet.Cells["B4"].PutValue("Furniture");
                sheet.Cells["C4"].PutValue(250);

                sheet.Cells["A5"].PutValue("Chair");
                sheet.Cells["B5"].PutValue("Furniture");
                sheet.Cells["C5"].PutValue(150);

                // Convert the plain range into a structured table (ListObject)
                int tableIndex = sheet.ListObjects.Add("A1", "C5", true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Set advanced table features
                table.DisplayName = "ProductTable";
                table.TableStyleType = TableStyleType.TableStyleMedium9;
                table.ShowTotals = true;
                table.ListColumns[2].TotalsCalculation = TotalsCalculation.Sum;

                // Apply style to the underlying range
                table.ApplyStyleToRange();

                // Save the workbook
                string outputPath = "ConvertedRangeToTable.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
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
            ConvertRangeToTableDemo.Run();
        }
    }
}
