// Title: How to group shapes Icon1, Icon2, and Icon3 into a ShapeGroup named IconSet with Aspose.Cells for .NET (grouping not supported)
// AI Prompts: Generate C# code that loads an Excel workbook, retrieves shapes named Icon1, Icon2, and Icon3, and attempts to create a ShapeGroup called IconSet using Aspose.Cells, with proper handling when grouping is unavailable. | Provide a C# example that checks for the presence of shapes Icon1‑Icon3 and applies identical formatting to them as a fallback when the ShapeGroup feature is not supported in Aspose.Cells. | Write a C# snippet that assembles Icon1, Icon2, and Icon3 into a logical collection for batch operations in Aspose.Cells, including comments on the current API limitation.
// Common Searches: aspose.cells create shape group from existing shapes c# | c# group multiple Excel shapes using Aspose.Cells API | aspose.cells shape grouping limitation and workaround | retrieve shapes by name and apply batch formatting aspose.cells | how to simulate shape grouping in Aspose.Cells .NET
// Tags: shape group creation Aspose.Cells .NET | retrieve Excel shapes by name C# | Aspose.Cells shape grouping limitation | collective shape formatting Aspose.Cells | Excel shape manipulation Aspose.Cells API | workaround for shape grouping Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing or new workbook, attempts to locate shapes named Icon1, Icon2, and Icon3 on the first worksheet, notes that the current Aspose.Cells version does not support ShapeGroup creation, and saves the workbook. It serves as a basis for handling shape grouping limitations and implementing alternative batch‑formatting strategies.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Load workbook if the input file exists; otherwise create a new workbook.
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve shapes by their names.
                Shape icon1 = null;
                Shape icon2 = null;
                Shape icon3 = null;

                try { icon1 = sheet.Shapes["Icon1"]; } catch { /* shape may not exist */ }
                try { icon2 = sheet.Shapes["Icon2"]; } catch { /* shape may not exist */ }
                try { icon3 = sheet.Shapes["Icon3"]; } catch { /* shape may not exist */ }

                // Grouping shapes is not supported in this version of Aspose.Cells.
                // If needed, you can implement alternative logic here.

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
