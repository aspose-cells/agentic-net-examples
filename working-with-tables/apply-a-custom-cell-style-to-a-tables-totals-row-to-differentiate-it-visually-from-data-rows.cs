using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class ApplyCustomStyleToTotalsRow
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data (header + 3 data rows)
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Price");
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("Orange");
                cells["B3"].PutValue(15);
                cells["A4"].PutValue("Banana");
                cells["B4"].PutValue(12);

                // Add a table that includes the data range and enable totals row
                int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.ShowTotals = true;

                // Set totals calculation for the Price column (second column)
                ListColumn priceColumn = table.ListColumns[1];
                priceColumn.TotalsCalculation = TotalsCalculation.Sum;
                priceColumn.TotalsRowLabel = "Grand Total";

                // ------------------------------------------------------------
                // Create a custom table style that will be applied to the totals row
                // ------------------------------------------------------------
                const string customStyleName = "MyTotalsStyle";

                // Add a new table style to the workbook's TableStyles collection
                TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;
                int styleIdx = tableStyles.AddTableStyle(customStyleName);
                TableStyle customTableStyle = tableStyles[styleIdx];

                // Define the style for the totals row
                Style totalsRowStyle = workbook.CreateStyle();
                totalsRowStyle.Pattern = BackgroundType.Solid;
                totalsRowStyle.ForegroundColor = Color.LightGoldenrodYellow; // background color
                totalsRowStyle.Font.IsBold = true;                           // bold font
                totalsRowStyle.Font.Color = Color.DarkBlue;                  // font color

                // Associate the style with the TotalRow element of the custom table style
                TableStyleElementCollection elements = customTableStyle.TableStyleElements;
                int elementIdx = elements.Add(TableStyleElementType.TotalRow);
                TableStyleElement totalRowElement = elements[elementIdx];
                totalRowElement.SetElementStyle(totalsRowStyle);

                // Apply the custom table style to the table
                table.TableStyleName = customStyleName;

                // ------------------------------------------------------------
                // Save the workbook
                // ------------------------------------------------------------
                string outputPath = "TableWithCustomTotalsRowStyle.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyCustomStyleToTotalsRow.Run();
        }
    }
}