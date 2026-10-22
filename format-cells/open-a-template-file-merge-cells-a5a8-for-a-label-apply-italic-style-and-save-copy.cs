// Title: Merge cells A5:A8, apply italic font style, and save a new workbook from a template using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells merge range A5:A8 and set italic font | How to apply italic style to merged cells with Aspose.Cells .NET | Saving a modified Excel template as a new file using Aspose.Cells in C# | Create StyleFlag for italic text on a specific range in Aspose.Cells | Merge cells vertically and format text in an Aspose.Cells workbook
// Tags: merge cells range Aspose.Cells C# | apply italic style to range Aspose.Cells | save workbook as new file Aspose.Cells | load template workbook Aspose.Cells | styleflag fontitalic Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads a template.xlsx workbook, merges cells A5 through A8 on the first worksheet, creates an italic font style using a Style and StyleFlag, applies the style to the merged range, and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "template.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file \"{templatePath}\" not found.");
                return;
            }

            // Load the template workbook
            Workbook workbook = new Workbook(templatePath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells A5:A8 (rows 4-7, column 0)
            sheet.Cells.Merge(4, 0, 4, 1);

            // Create an italic font style
            Style italicStyle = workbook.CreateStyle();
            italicStyle.Font.IsItalic = true;

            // Define which style attributes to apply
            StyleFlag styleFlag = new StyleFlag
            {
                FontItalic = true
            };

            // Apply the italic style to the merged range
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("A5:A8");
            mergedRange.ApplyStyle(italicStyle, styleFlag);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
