// Title: Show a table’s header row and apply bold formatting to header cells with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook, sets ListObject.ShowHeaderRow = true, and applies a bold font style to the header row using Aspose.Cells. | Provide a reusable method that takes a worksheet and formats the first table’s header row as bold while ensuring the header row is visible with Aspose.Cells. | Write a snippet that creates a bold Style, applies it to the header range of a ListObject, and saves the workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# enable table header row and make header bold | How to set ShowHeaderRow true and apply bold style to Excel table header using Aspose.Cells | C# Aspose.Cells format ListObject header row with bold font | Make Excel table header visible and bold with Aspose.Cells .NET | Apply style to table header range Aspose.Cells C# example
// Tags: listobject showheaderrow aspocells | apply bold style to table header aspocells | excel table header formatting c# | styleflag fontbold aspocells | create range header row aspocells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables; // For ListObject
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

// Loads an existing workbook, ensures the first ListObject’s header row is visible, creates a bold font style, applies it to the header range, and saves the result to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables (ListObjects) found in the worksheet.");
                return;
            }

            // Get the first table
            ListObject table = sheet.ListObjects[0];
            table.ShowHeaderRow = true; // Ensure header row is visible

            // Create a range that represents the header row of the table
            int headerRow = table.StartRow;               // first row of the table (header)
            int startColumn = table.StartColumn;
            int columnCount = table.ListColumns.Count;   // correct way to get column count

            AsposeRange headerRange = sheet.Cells.CreateRange(headerRow, startColumn, 1, columnCount);

            // Define a bold font style
            Style boldStyle = workbook.CreateStyle();
            boldStyle.Font.IsBold = true;

            // Apply the style to the header range
            StyleFlag flag = new StyleFlag { FontBold = true, Font = true };
            headerRange.ApplyStyle(boldStyle, flag);

            // Ensure output directory exists
            try
            {
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
