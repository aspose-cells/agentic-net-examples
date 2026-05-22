using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class ListObjectHeaderStyleDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the table (including header)
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Price");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(2.5);
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(1.8);
                worksheet.Cells["A4"].PutValue("Cherry");
                worksheet.Cells["B4"].PutValue(3.2);

                // Add a ListObject (table) covering the data range
                int tableIndex = worksheet.ListObjects.Add("A1", "B4", true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Ensure the header row is visible
                table.ShowHeaderRow = true;

                // Create a custom style for the header row
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Pattern = BackgroundType.Solid;
                headerStyle.BackgroundColor = Color.LightGreen; // Desired background color
                headerStyle.Font.IsBold = true;                  // Optional: make header text bold

                // Create a new table style and set the HeaderRow element style
                string customStyleName = "MyHeaderStyle";
                TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;
                int styleIdx = tableStyles.AddTableStyle(customStyleName);
                TableStyle customTableStyle = tableStyles[styleIdx];

                TableStyleElementCollection elements = customTableStyle.TableStyleElements;
                int headerElementIdx = elements.Add(TableStyleElementType.HeaderRow);
                TableStyleElement headerElement = elements[headerElementIdx];
                headerElement.SetElementStyle(headerStyle);

                // Apply the custom table style to the list object
                table.TableStyleName = customStyleName;

                // Save the workbook
                string outputPath = "ListObjectHeaderStyleDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
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
            ListObjectHeaderStyleDemo.Run();
        }
    }
}