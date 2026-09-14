// Title: Add a cell‑value conditional formatting rule to an XLS workbook and export it as an MHT web archive using Aspose.Cells for .NET
// AI Prompts: Create a conditional formatting rule that colors cells with values greater than 50 light green, then save the worksheet as an MHTML file. | Load an existing .xls file, apply a greater‑than‑50 cell‑value format, and generate a .mht web archive with Aspose.Cells in C#.
// Common Searches: asp.net how to apply conditional formatting to a range and export to mhtml with Aspose.Cells | c# code to highlight cells greater than 50 in an xls file and save as mht | using Aspose.Cells to convert xls to mhtml while preserving conditional formats | conditional formatting example Aspose.Cells C# export workbook to web archive | generate MHTML web archive from workbook after applying cell value rule
// Tags: cell value conditional format Aspose.Cells | MHTML web archive creation Aspose.Cells | XLS workbook conversion to MHT Aspose.Cells | light green background style Aspose.Cells | range A1:A10 formatting Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLS workbook (or creates one with sample data), adds a conditional formatting rule that highlights cells A1:A10 with values over 50 using a light‑green background, and saves the result as an MHTML web archive (output.mht) via Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xls";
            const string outputFile = "output.mht";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputFile))
            {
                workbook = new Workbook(inputFile);
            }
            else
            {
                workbook = new Workbook();
                Worksheet wsNew = workbook.Worksheets[0];
                // Populate sample data (0,10,20,...,90) for demonstration
                for (int i = 0; i < 10; i++)
                {
                    wsNew.Cells[i, 0].PutValue(i * 10);
                }
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range A1:A10 (zero‑based indices)
            int firstRow = 0;
            int firstCol = 0;
            int totalRows = 10;
            int totalCols = 1;

            // Add a new ConditionalFormatting rule for the defined range
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex];
            cf.AddArea(new CellArea
            {
                StartRow = firstRow,
                StartColumn = firstCol,
                EndRow = firstRow + totalRows - 1,
                EndColumn = firstCol + totalCols - 1
            });

            // Create a condition: cell value > 50
            int conditionIdx = cf.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "50",
                ""); // second formula required by API (unused for this operator)

            // Retrieve the created condition
            FormatCondition condition = cf[conditionIdx];

            // Define the style to apply (light green background)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightGreen;
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

            // Save the workbook as an MHTML (web archive) file
            workbook.Save(outputFile, SaveFormat.MHtml);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
