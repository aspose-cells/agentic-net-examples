// Title: Set a custom paper size for a worksheet in Aspose.Cells for .NET using reflection or the SetCustomPaperSize method
// Common Searches: how to set a custom page size in Aspose.Cells .NET workbook | using reflection to assign PaperWidth and PaperHeight in Aspose.Cells PageSetup | Aspose.Cells SetCustomPaperSize method example C# | override read‑only PageSetup properties in Aspose.Cells | apply A4 dimensions to a worksheet with Aspose.Cells
// Tags: custom paper size PageSetup Aspose.Cells | set PaperWidth PaperHeight via reflection C# | SetCustomPaperSize method Aspose.Cells | override read‑only PageSetup properties | A4 dimensions workbook Aspose.Cells | apply custom page size to worksheet .NET

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsCustomPaperSize
{
    // The example loads or creates a workbook, accesses the first worksheet, sets its PageSetup to a custom paper size, and assigns PaperWidth and PaperHeight using reflection; if the properties are read‑only it falls back to the SetCustomPaperSize method, then saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Load an existing workbook if it exists; otherwise create a new one
                Workbook workbook;
                string inputPath = "input.xlsx";

                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Custom paper size in points (1 point = 1/72 inch)
                double paperWidth = 595.0;   // A4 width
                double paperHeight = 842.0;  // A4 height

                // Apply custom paper size
                PageSetup pageSetup = sheet.PageSetup;
                pageSetup.PaperSize = PaperSizeType.Custom;

                // Attempt to set PaperWidth and PaperHeight via reflection (handles read‑only properties in some versions)
                PropertyInfo widthProp = pageSetup.GetType().GetProperty("PaperWidth");
                PropertyInfo heightProp = pageSetup.GetType().GetProperty("PaperHeight");

                if (widthProp != null && widthProp.CanWrite && heightProp != null && heightProp.CanWrite)
                {
                    widthProp.SetValue(pageSetup, paperWidth);
                    heightProp.SetValue(pageSetup, paperHeight);
                }
                else
                {
                    // Fallback: use SetCustomPaperSize method if available
                    MethodInfo setCustomMethod = pageSetup.GetType().GetMethod("SetCustomPaperSize");
                    if (setCustomMethod != null)
                    {
                        setCustomMethod.Invoke(pageSetup, new object[] { paperWidth, paperHeight });
                    }
                }

                // Save the workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
