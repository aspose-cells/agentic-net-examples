// Title: Apply a light‑gray background style to a nested list range using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a Style with a solid light‑gray fill, configures a StyleFlag for CellShading, and applies the style to the range A1:A5 on the first worksheet. | Generate a C# example that populates column A with a simple hierarchical list, applies the background style to the defined range, and saves the workbook as NestedListWithIndentation.xlsx.
// Common Searches: asp.net apply solid fill style to a range with Aspose.Cells | c# aspose.cells StyleFlag cell shading example | how to style a column range with background color in Aspose.Cells | create nested list in Excel and apply custom style using Aspose.Cells C# | aspose.cells save workbook after applying style to range A1:A5
// Tags: apply background style to range Aspose.Cells C# | StyleFlag cell shading Aspose.Cells | nested list range styling Aspose.Cells | solid fill style creation Aspose.Cells | save styled workbook Aspose.Cells C#

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsIndentationExample
{
    // The program creates a new workbook, fills cells A1‑A5 with a simple hierarchical list, defines a light‑gray solid fill style, uses a StyleFlag to apply cell shading, applies the style to the range A1:A5, and saves the file as NestedListWithIndentation.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Prepare a sample nested list in column A
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Item 1");
                sheet.Cells["A2"].PutValue("  Subitem 1.1"); // visual indent for demo
                sheet.Cells["A3"].PutValue("  Subitem 1.2");
                sheet.Cells["A4"].PutValue("Item 2");
                sheet.Cells["A5"].PutValue("  Subitem 2.1");

                // -------------------------------------------------
                // Apply a simple style to the range (indentation not supported in this version)
                // -------------------------------------------------
                int startRow = 0;      // Row index for A1
                int startColumn = 0;   // Column index for A
                int totalRows = 5;     // Number of rows in the list
                int totalColumns = 1;  // Only column A

                // Create a style (customize other attributes if needed)
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.LightGray;
                style.Pattern = BackgroundType.Solid;

                // Define which style attributes will be applied
                StyleFlag styleFlag = new StyleFlag
                {
                    CellShading = true // Apply background color
                };

                // Apply the style to the defined range
                Aspose.Cells.Range listRange = sheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);
                listRange.ApplyStyle(style, styleFlag);

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "NestedListWithIndentation.xlsx";

                // Ensure the directory exists (if a directory part is present)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
