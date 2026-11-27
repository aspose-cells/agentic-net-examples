// Title: Validate that QuotePrefix remains set on cells with mixed formatting after applying a StyleFlag in Aspose.Cells for .NET
// Common Searches: how to retain leading quote in a cell after using StyleFlag in Aspose.Cells C# | Aspose.Cells apply font color to a range without changing QuotePrefix property | preserve quote prefix when updating cell style with StyleFlag in .NET | validate QuotePrefix flag remains true after style changes in Aspose.Cells workbook | C# example for applying StyleFlag while keeping mixed formatting cells unchanged
// Tags: Aspose.Cells StyleFlag font update preserving QuotePrefix | maintain leading apostrophe when applying range style in .NET | mixed formatting cells QuotePrefix retention Aspose.Cells | apply style to Excel range without altering quote prefix flag | C# workbook style modification using StyleFlag with QuotePrefix safety

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a workbook, sets QuotePrefix on cell A1 while leaving B1 without it, defines a style that changes font color to red and makes text bold, configures a StyleFlag to modify only FontColor and FontBold, applies this style to the range A1:B1, and then verifies that A1 still has QuotePrefix true and B1 remains false, outputting the results and saving the file as QuotePrefixValidation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // -------------------------------------------------
            // 1. Prepare cells with mixed formatting
            // -------------------------------------------------

            // Cell A1: value that should retain a leading quote prefix
            Cell cellA1 = cells["A1"];
            cellA1.PutValue("12345"); // put a plain string value
            Style styleWithQuote = workbook.CreateStyle();
            styleWithQuote.QuotePrefix = true;               // enable quote prefix
            styleWithQuote.Font.Name = "Calibri";
            styleWithQuote.Font.Size = 11;
            cellA1.SetStyle(styleWithQuote);

            // Cell B1: normal cell without a quote prefix
            Cell cellB1 = cells["B1"];
            cellB1.PutValue("Normal Text");
            Style normalStyle = workbook.CreateStyle();
            normalStyle.Font.Name = "Calibri";
            normalStyle.Font.Size = 11;
            cellB1.SetStyle(normalStyle);

            // -------------------------------------------------
            // 2. Define a new style that will be applied via StyleFlag
            // -------------------------------------------------
            Style updateStyle = workbook.CreateStyle();
            updateStyle.Font.Color = Color.Red;   // change font color to red
            updateStyle.Font.IsBold = true;      // make text bold

            // -------------------------------------------------
            // 3. Configure StyleFlag – we only want to change FontColor and FontBold,
            //    NOT the QuotePrefix flag.
            // -------------------------------------------------
            StyleFlag flag = new StyleFlag
            {
                FontColor = true,
                FontBold = true
                // QuotePrefix remains false (default) so it will not be altered
            };

            // Apply the style to the range that contains both cells
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1:B1");
            range.ApplyStyle(updateStyle, flag);

            // -------------------------------------------------
            // 4. Validate that the QuotePrefix property is preserved
            // -------------------------------------------------
            bool a1QuotePrefix = cellA1.GetStyle().QuotePrefix;
            bool b1QuotePrefix = cellB1.GetStyle().QuotePrefix;

            Console.WriteLine($"Cell A1 QuotePrefix after StyleFlag update: {a1QuotePrefix}");
            Console.WriteLine($"Cell B1 QuotePrefix after StyleFlag update: {b1QuotePrefix}");

            // Simple assertions
            if (a1QuotePrefix)
                Console.WriteLine("SUCCESS: QuotePrefix retained on cell A1.");
            else
                Console.WriteLine("ERROR: QuotePrefix was lost on cell A1.");

            if (!b1QuotePrefix)
                Console.WriteLine("SUCCESS: Cell B1 correctly has no QuotePrefix.");
            else
                Console.WriteLine("ERROR: QuotePrefix incorrectly set on cell B1.");

            // Save the workbook for manual inspection
            workbook.Save("QuotePrefixValidation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
