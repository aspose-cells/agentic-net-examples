// Title: Use Aspose.Cells CopyStyle to transfer locale‑specific date formatting from a template cell to generated date cells in C#
// AI Prompts: Copy the date format from cell A1 of a template workbook and apply it to a range of new date cells using Aspose.Cells CopyStyle in .NET. | Transfer a locale‑specific date style from a source worksheet to multiple destination cells after inserting DateTime values with Aspose.Cells.
// Common Searches: Aspose.Cells copy date cell style from template workbook to another workbook in C# | C# how to inherit locale specific date format using CopyStyle method Aspose.Cells | apply template date formatting to generated dates with Aspose.Cells .NET | example of copying cell style between ranges for date formatting in Aspose.Cells
// Tags: CopyStyle date formatting Aspose.Cells | inherit locale date style C# | transfer cell style between workbooks Aspose.Cells | apply template formatting to generated dates .NET | date cell style copy range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCopyDateStyleDemo
{
    // The example loads a template workbook, extracts the date style from cell A1, writes three DateTime values to a new workbook, copies the template's style to the range A1:C1 using CopyStyle, and saves the result as Result.xlsx.
    class Program
    {
        static void Main()
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            try
            {
                // Verify that the template file exists before loading
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file \"{templatePath}\" not found.");
                    return;
                }

                // Load the template workbook that contains the desired date format in cell A1
                Workbook templateWorkbook = new Workbook(templatePath);
                Worksheet templateSheet = templateWorkbook.Worksheets[0];

                // Define the source range that holds the date style (single cell A1)
                Aspose.Cells.Range sourceDateRange = templateSheet.Cells.CreateRange("A1");

                // Create a new workbook where dates will be generated
                Workbook resultWorkbook = new Workbook();
                Worksheet resultSheet = resultWorkbook.Worksheets[0];

                // Populate some date values in the result sheet
                resultSheet.Cells["A1"].PutValue(DateTime.Now);
                resultSheet.Cells["B1"].PutValue(DateTime.Now.AddDays(1));
                resultSheet.Cells["C1"].PutValue(DateTime.Now.AddDays(2));

                // Define the destination range that should inherit the date formatting
                Aspose.Cells.Range destinationDateRange = resultSheet.Cells.CreateRange("A1:C1");

                // Copy the style (including date format) from the template range to the destination range
                destinationDateRange.CopyStyle(sourceDateRange);

                // Save the resulting workbook
                resultWorkbook.Save(resultPath);
                Console.WriteLine($"Result workbook saved to \"{resultPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
