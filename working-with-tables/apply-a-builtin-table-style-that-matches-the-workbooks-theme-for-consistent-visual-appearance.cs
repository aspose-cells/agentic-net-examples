using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class ApplyBuiltinTableStyleDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the table
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Quantity");
                worksheet.Cells["C1"].PutValue("Price");

                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["C2"].PutValue(0.5);

                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["C3"].PutValue(0.3);

                worksheet.Cells["A4"].PutValue("Cherry");
                worksheet.Cells["B4"].PutValue(15);
                worksheet.Cells["C4"].PutValue(0.8);

                // Add a ListObject (table) that covers the data range
                int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 2, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Get the collection of built‑in table styles
                TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;

                // Choose a built‑in style that aligns with the workbook's theme.
                TableStyle builtinStyle = tableStyles.GetBuiltinTableStyle(TableStyleType.TableStyleMedium2);

                // Apply the built‑in style to the table
                table.TableStyleName = builtinStyle.Name;

                // Optional: enable additional style features
                table.ShowTableStyleFirstColumn = true;
                table.ShowTableStyleLastColumn = true;
                table.ShowTableStyleRowStripes = true;
                table.ShowTableStyleColumnStripes = true;

                // Save the workbook
                string outputPath = "AppliedBuiltinTableStyle.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}