// Title: Create an Excel table with a styled header row and custom display name using Aspose.Cells for .NET
// AI Prompts: Write C# code that builds a new Workbook, inserts sample product data, applies a bold light‑blue style to the header cells, adds a ListObject covering the data range, assigns a custom DisplayName to the table, and saves the workbook to a specified file. | Show how to verify that the target folder exists and create it if necessary before calling Workbook.Save in Aspose.Cells.
// Common Searches: Aspose.Cells C# create Excel table with formatted header | how to set ListObject DisplayName in Aspose.Cells .NET | preserve header cell style after refreshing data in Aspose.Cells query table | ensure output directory exists before saving workbook Aspose.Cells
// Tags: Aspose.Cells create ListObject with header formatting | C# set ListObject display name Aspose.Cells | preserve formatting on data refresh Aspose.Cells | validate output folder before saving workbook C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example demonstrates how to generate a new Workbook, populate it with product data, apply a bold light‑blue style to the header row, create a ListObject (Excel table) over the range, assign the display name "ProductsTable", ensure the destination directory exists, and finally save the file as QueryTablePreserveFormattingDemo.xlsx using Aspose.Cells for .NET.
class QueryTablePreserveFormattingDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data that will be used by the table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Laptop");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Phone");
            sheet.Cells["B3"].PutValue(800);

            // Apply a style to the header row (A1:B1)
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.LightBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            sheet.Cells["A1"].SetStyle(headerStyle);
            sheet.Cells["B1"].SetStyle(headerStyle);

            // Create a ListObject (Excel table) covering the data range A1:B3
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = sheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            // Set the display name of the table (Aspose.Cells uses DisplayName instead of Name)
            table.DisplayName = "ProductsTable";

            // Output the current display name to verify
            Console.WriteLine("Table display name is set to: " + table.DisplayName);

            // Save the workbook
            string outputPath = "QueryTablePreserveFormattingDemo.xlsx";

            // Ensure the directory exists if a path is provided
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        QueryTablePreserveFormattingDemo.Run();
    }
}
