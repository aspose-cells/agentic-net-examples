// Title: Copy a Worksheet with Formatting to Another Workbook using Aspose.Cells for .NET
// Description: Shows how to load or create a source workbook, apply a style to a cell, and use Worksheet.Copy to duplicate the sheet—including data, fonts, colors, borders, and conditional formats—into a new workbook, then save the result.
// Keywords: Aspose.Cells copy worksheet | duplicate sheet with formatting | preserve Excel styles C# | Worksheet.Copy method | transfer sheet between workbooks | C# Excel copy formatting | Aspose.Cells .NET example
// Common Searches: Aspose.Cells copy worksheet with formatting | C# copy Excel sheet preserving styles | How to duplicate a sheet to another workbook Aspose | Worksheet.Copy preserving cell formatting | Copy Excel worksheet to new file using Aspose.Cells
// Developer Intent: Duplicate a worksheet from one Excel file to another while keeping all formatting intact.
// Use Cases: Create a branded template workbook and reuse its styled sheet across multiple reports. | Migrate data from an existing analysis file into a fresh workbook without losing any visual formatting. | Generate personalized Excel files by inserting a pre‑styled worksheet as a new tab for each recipient.
// AI Prompts: Provide C# code that copies a worksheet with all formatting from one workbook to another using Aspose.Cells. | Explain how Worksheet.Copy preserves conditional formatting, data validation, and cell styles in Aspose.Cells for .NET. | Show how to add a copied worksheet as a new tab when the destination workbook already contains other sheets.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to load or create a source workbook, apply a style to a cell, and use Worksheet.Copy to duplicate the sheet—including data, fonts, colors, borders, and conditional formats—into a new workbook, then save the result.
    public class CopyWorksheetWithStylesDemo
    {
        public static void Run()
        {
            const string sourcePath = "source.xlsx";
            const string destPath = "destination.xlsx";

            try
            {
                // Ensure source workbook exists; create a simple one if missing
                Workbook sourceWorkbook;
                if (File.Exists(sourcePath))
                {
                    sourceWorkbook = new Workbook(sourcePath);
                }
                else
                {
                    sourceWorkbook = new Workbook();
                    Worksheet ws = sourceWorkbook.Worksheets[0];
                    ws.Name = "SampleSheet";
                    ws.Cells["A1"].PutValue("Hello");

                    // Apply style to the cell
                    Style style = ws.Cells["A1"].GetStyle();
                    style.Font.Color = Color.Blue;
                    ws.Cells["A1"].SetStyle(style);

                    sourceWorkbook.Save(sourcePath);
                }

                // Create an empty destination workbook
                Workbook destWorkbook = new Workbook();

                // Get the source and destination worksheets
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
                Worksheet destSheet = destWorkbook.Worksheets[0];

                // Copy the source worksheet (data + styles) to the destination worksheet
                destSheet.Copy(sourceSheet);

                // Save the destination workbook
                destWorkbook.Save(destPath);
                Console.WriteLine($"Worksheet copied successfully to '{destPath}'.");
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
            CopyWorksheetWithStylesDemo.Run();
        }
    }
}
