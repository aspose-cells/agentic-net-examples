// Title: Change the font color of a pivot table header cell to black using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a pivot table and sets the first header cell's font color to black. | Show how to apply a custom style with a black font to a pivot table header cell in a .NET workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# set pivot table header font color to black | how to format pivot table header cell font color using Aspose.Cells .NET | apply custom style to pivot field header in Excel with Aspose.Cells C# | change pivot table header text color programmatically Aspose.Cells | C# Aspose.Cells example for styling pivot table header cell
// Tags: Aspose.Cells pivot table header font styling | C# set pivot field header color Aspose | apply style to pivot table header cell .NET | format pivot table header text color Excel | Aspose.Cells workbook style black font

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable and PivotFieldType

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, builds a pivot table, defines a style with a black font, applies that style to the pivot table's first header cell, and saves the file as an Excel workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Bike");
                sheet.Cells["B2"].PutValue(1000);
                sheet.Cells["A3"].PutValue("Car");
                sheet.Cells["B3"].PutValue(2000);
                sheet.Cells["A4"].PutValue("Bike");
                sheet.Cells["B4"].PutValue(1500);
                sheet.Cells["A5"].PutValue("Car");
                sheet.Cells["B5"].PutValue(3000);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table (Product as row field, Sales as data field)
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column

                // Calculate the pivot table so that it is populated
                pivotTable.CalculateData();

                // Create a style with black font color
                Style blackFontStyle = workbook.CreateStyle();
                blackFontStyle.Font.Color = Color.Black;

                // Apply the style to the pivot table header cell (first header cell at row 0, column 0)
                pivotTable.Format(0, 0, blackFontStyle);

                // Define output file path
                string outputPath = "PivotHeaderBlackFont.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
