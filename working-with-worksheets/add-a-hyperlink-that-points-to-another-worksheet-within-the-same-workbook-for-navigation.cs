// Title: Add an internal hyperlink to a different worksheet in an Aspose.Cells workbook with C#
// Common Searches: Aspose.Cells C# create hyperlink that jumps to another sheet in the same workbook | how to link Sheet1 cell to Sheet2 using Aspose.Cells API in C# | C# internal worksheet hyperlink example with Aspose.Cells Save as .xlsx | navigate between worksheets programmatically with Aspose.Cells hyperlink | Aspose.Cells add internal link to cell A1 of another worksheet in C#
// Tags: Aspose.Cells internal worksheet hyperlink C# | add hyperlink to another sheet Aspose.Cells | hyperlink between worksheets Aspose.Cells API | C# Aspose.Cells navigation link | create internal link in .xlsx using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

namespace HyperlinkExample
{
    // The sample creates a new workbook, adds two worksheets, inserts a hyperlink in Sheet1!A1 that points to Sheet2!A1, and saves the file as HyperlinkExample.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (source)
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Sheet1";

                // Add a second worksheet (target)
                int targetIndex = workbook.Worksheets.Add();
                Worksheet targetSheet = workbook.Worksheets[targetIndex];
                targetSheet.Name = "Sheet2";

                // Put display text in the source cell
                Cell sourceCell = sourceSheet.Cells["A1"];
                sourceCell.PutValue("Go to Sheet2");

                // Add an internal hyperlink that points to Sheet2!A1
                sourceSheet.Hyperlinks.Add(sourceCell.Row, sourceCell.Column, 1, 1, "Sheet2!A1");

                // Save the workbook
                string outputPath = "HyperlinkExample.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
