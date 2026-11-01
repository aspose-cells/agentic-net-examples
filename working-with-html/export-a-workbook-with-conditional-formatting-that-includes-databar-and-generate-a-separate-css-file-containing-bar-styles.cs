// Title: Create an Excel .xlsx workbook with a LightGreen DataBar conditional format and generate a matching CSS stylesheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to add a LightGreen DataBar conditional format to the range A2:A5 and save the workbook as an .xlsx file. | Generate a .css file in C# that defines a .databar class with a linear‑gradient mirroring the LightGreen DataBar, using a CSS custom property for the bar length. | Extend the example to apply different DataBar colors to multiple columns and create separate CSS classes for each, then export the workbook and all CSS files.
// Common Searches: how to add a DataBar conditional format with Aspose.Cells C# and export to XLSX | generate CSS stylesheet that mimics Excel DataBar colors using .NET | Aspose.Cells create workbook with DataBar and write external CSS file | C# export Excel file with conditional formatting and separate CSS for DataBar | save DataBar conditional formatting as CSS variable using Aspose.Cells
// Tags: Aspose.Cells DataBar conditional formatting export | C# generate CSS for Excel DataBar | export workbook as XLSX with conditional formatting | create external CSS stylesheet from Aspose.Cells | DataBar visual style CSS linear gradient

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The sample program creates a new workbook, fills column A with sample scores, adds a LightGreen DataBar conditional format to cells A2:A5, saves the file as DataBarWorkbook.xlsx, and writes a databar.css stylesheet containing a .databar class that uses a linear‑gradient to replicate the DataBar appearance.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A
            sheet.Cells["A1"].PutValue("Score");
            sheet.Cells["A2"].PutValue(30);
            sheet.Cells["A3"].PutValue(70);
            sheet.Cells["A4"].PutValue(50);
            sheet.Cells["A5"].PutValue(90);

            // Add a conditional formatting rule to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex]; // ConditionalFormatting instance

            // Define the range A2:A5 for the DataBar
            CellArea area = new CellArea
            {
                StartRow = 1,   // Row 2 (zero‑based)
                StartColumn = 0, // Column A
                EndRow = 4,     // Row 5
                EndColumn = 0   // Column A
            };
            cf.AddArea(area);

            // Create a DataBar condition
            int conditionIndex = cf.AddCondition(FormatConditionType.DataBar);
            FormatCondition dataBarCondition = cf[conditionIndex];

            // Configure DataBar appearance
            dataBarCondition.DataBar.MinLength = 0;               // Minimum length as percent
            dataBarCondition.DataBar.MaxLength = 100;             // Maximum length as percent
            dataBarCondition.DataBar.Color = Color.LightGreen;   // Bar color
            dataBarCondition.DataBar.ShowValue = true;            // Show the cell value inside the bar

            // Save the workbook to an Excel file
            string workbookPath = "DataBarWorkbook.xlsx";
            try
            {
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to save workbook: {ex.Message}");
            }

            // -----------------------------------------------------------------
            // Generate a CSS file that can be used to mimic the DataBar style
            // -----------------------------------------------------------------
            string css = @"
.databar {
    display: inline-block;
    height: 16px;
    /* The variable --value should be set to the percentage (e.g., 70%) */
    background: linear-gradient(to right, #90EE90 0%, #90EE90 var(--value), #fff var(--value), #fff 100%);
    border: 1px solid #ccc;
    box-sizing: border-box;
}
";
            string cssPath = "databar.css";

            try
            {
                File.WriteAllText(cssPath, css);
                Console.WriteLine($"CSS file written to {Path.GetFullPath(cssPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to write CSS file: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
