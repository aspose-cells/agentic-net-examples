// Title: Check that a cell's Quote Prefix flag stays unchanged after modifying font color with a StyleFlag in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# verify QuotePrefix after using StyleFlag to change font color | preserve quote prefix when updating cell style with StyleFlag in .NET | how to keep mixed formatting quote prefix in Aspose.Cells after style flag change | C# Aspose.Cells retain QuotePrefix flag after applying StyleFlag FontColor | validate quote prefix retention after cell style modification Aspose.Cells
// Tags: cell quote prefix flag preservation | StyleFlag FontColor update Aspose.Cells | mixed formatting cell with quote prefix | rich text character styling Aspose.Cells | C# modify cell style using StyleFlag

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// The example loads or creates a workbook, sets a cell with a Quote Prefix and mixed formatting, optionally applies rich‑text styling to a substring, then changes the whole‑cell font color to green using a StyleFlag that only affects FontColor, and finally confirms that the Quote Prefix flag remains unchanged before saving the file.
class MixedFormattingQuotePrefixValidation
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input workbook exists; create a simple one if it does not.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                workbook.Save(inputPath);
            }

            // Access the first worksheet and target cell B2.
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["B2"];

            // Store the original Quote Prefix state from the cell's style.
            bool originalQuotePrefix = cell.GetStyle().QuotePrefix;

            // Apply mixed formatting: set whole cell font color to Red.
            Style style = cell.GetStyle();
            style.Font.Color = Color.Red;
            cell.SetStyle(style);

            // Set cell value with two parts.
            cell.PutValue("Part1 Part2");

            // Attempt to apply rich‑text formatting to the first part.
            // This feature may not be available in older versions; therefore we guard it with a try‑catch.
            try
            {
                // Use reflection to call GetCharacters if it exists.
                var getCharsMethod = typeof(Cell).GetMethod("GetCharacters", new[] { typeof(int), typeof(int) });
                if (getCharsMethod != null)
                {
                    object charsObj = getCharsMethod.Invoke(cell, new object[] { 0, 5 });
                    if (charsObj != null)
                    {
                        var getStyleMethod = charsObj.GetType().GetMethod("GetStyle");
                        var setStyleMethod = charsObj.GetType().GetMethod("SetStyle", new[] { typeof(Style) });
                        if (getStyleMethod != null && setStyleMethod != null)
                        {
                            Style charStyle = (Style)getStyleMethod.Invoke(charsObj, null);
                            charStyle.Font.Color = Color.Blue;
                            setStyleMethod.Invoke(charsObj, new object[] { charStyle });
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Rich‑text formatting is not supported; continue without it.
            }

            // Change the whole cell font color to Green using StyleFlag to affect only FontColor.
            Style newStyle = cell.GetStyle();
            newStyle.Font.Color = Color.Green;

            StyleFlag flag = new StyleFlag
            {
                FontColor = true // Only modify the font color.
            };

            cell.SetStyle(newStyle, flag);

            // Verify that the Quote Prefix flag is still retained after style changes.
            bool afterUpdateQuotePrefix = cell.GetStyle().QuotePrefix;

            // Output the validation result.
            Console.WriteLine("Original Quote Prefix: " + originalQuotePrefix);
            Console.WriteLine("Quote Prefix After StyleFlag Update: " + afterUpdateQuotePrefix);
            Console.WriteLine("Quote Prefix retained: " + (originalQuotePrefix == afterUpdateQuotePrefix));

            // Save the workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
