// Title: Open an Excel template, merge cells A5:A8, apply italic font style, and save a new workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to merge a range of cells and set italic font | How to open an Excel template, merge cells A5 to A8 and save a new file with Aspose.Cells | Applying italic style to merged cells in an existing workbook using Aspose.Cells for .NET | Aspose.Cells C# code to merge cells and use StyleFlag for font formatting | Save a copy of a modified Excel workbook after merging cells with Aspose.Cells
// Tags: Aspose.Cells merge cell range C# | apply italic font style Aspose.Cells | open template workbook Aspose.Cells | save modified workbook Aspose.Cells | StyleFlag font italic Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel template (or creates a new workbook if the template is missing), merges cells A5 through A8 on the first worksheet, applies an italic font style using a StyleFlag, and saves the result as a separate output file.
class Program
{
    static void Main()
    {
        // Paths for the template and output workbooks
        string templatePath = "Template.xlsx";
        string outputPath = "Output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing template if it exists; otherwise create a new workbook
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range A5:A8 using Aspose.Cells.Range to avoid ambiguity with System.Range
            Aspose.Cells.Range labelRange = sheet.Cells.CreateRange("A5", "A8");

            // Merge the cells in the range
            labelRange.Merge();

            // Create a style with italic font
            Style italicStyle = workbook.CreateStyle();
            italicStyle.Font.IsItalic = true;

            // Specify that only the italic font attribute should be applied
            StyleFlag styleFlag = new StyleFlag
            {
                FontItalic = true
            };

            // Apply the italic style to the merged range
            labelRange.ApplyStyle(italicStyle, styleFlag);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
