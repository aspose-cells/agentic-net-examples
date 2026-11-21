// Title: Apply a template row’s background color to multiple rows using SetStyle in Aspose.Cells C#
// Common Searches: how to copy a row’s background color to other rows with Aspose.Cells in C# | C# Aspose.Cells set style for a range of rows based on a template row | inherit row shading from a template worksheet using Aspose.Cells | apply template row formatting to generated rows in an Excel file with Aspose.Cells | preserve background color when programmatically adding rows in Aspose.Cells .NET
// Tags: set row style Aspose.Cells C# | background color inheritance from template row | template row style propagation | Aspose.Cells style propagation Excel | programmatic row formatting .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a template workbook, extracts the style (including background color) from the second row, and uses SetStyle to apply that style to rows 3 through 11 before saving the result as Result.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            try
            {
                // Verify that the template file exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook
                var workbook = new Workbook(templatePath);
                var worksheet = workbook.Worksheets[0];

                // Row index (0‑based) that contains the desired background shading
                int templateRowIndex = 1; // second row in the sheet

                // Define the range of rows that will receive the style
                int startRowIndex = 2; // first row to receive the style
                int endRowIndex = 10;  // last row to receive the style

                // Get the style from the template row
                var sourceRow = worksheet.Cells.Rows[templateRowIndex];
                Style sourceStyle = sourceRow.GetStyle();

                // Apply the source row style to each target row
                for (int row = startRowIndex; row <= endRowIndex; row++)
                {
                    // Ensure the target row exists (Aspose.Cells creates rows on demand)
                    var targetRow = worksheet.Cells.Rows[row];
                    targetRow.SetStyle(sourceStyle);
                }

                // Save the modified workbook
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to {resultPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
