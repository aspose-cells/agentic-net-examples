// Title: Show ListObject header row and apply a LightBlue solid background style in Aspose.Cells for .NET
// AI Prompts: Create a ListObject table, enable its header row, and set the header background to LightBlue using a solid fill style in a C# Aspose.Cells workbook. | Apply a custom style with a LightBlue solid fill to the first row of a ListObject and save the workbook as an .xlsx file using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to display the header row of a ListObject table | set background color for ListObject header in Aspose.Cells .NET | apply solid fill style to table header using Aspose.Cells API | C# code to style ListObject header row with custom color in Excel workbook
// Tags: listobject showheaderrow aspnet | listobject header style aspnet cells | solid fill background style aspnet cells | c# aspnet cells table header formatting | save styled listobject workbook .xlsx

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, adds a ListObject covering cells A1:B3, enables the header row, defines a LightBlue solid background style, applies this style to the header row range, and saves the workbook as ListObjectWithHeaderStyle.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(30);

            // Define the range that will become a ListObject (table)
            int firstRow = 0;        // zero‑based index for row 1
            int firstColumn = 0;     // column A
            int totalRows = 3;       // rows 1‑3
            int totalColumns = 2;    // columns A‑B

            // Add a ListObject to the worksheet
            int listIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);
            ListObject list = sheet.ListObjects[listIndex];

            // Ensure the header row is shown
            list.ShowHeaderRow = true;

            // Create a style for the header background
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.LightBlue;
            headerStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the header row range (first row of the table)
            AsposeRange headerRange = sheet.Cells.CreateRange(firstRow, firstColumn, 1, totalColumns);
            StyleFlag flag = new StyleFlag { All = true };
            headerRange.ApplyStyle(headerStyle, flag);

            // Save the workbook
            string outputPath = "ListObjectWithHeaderStyle.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
