// Title: Hide a specific named range from the Name Manager using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide a named range from the Name Manager | Set IsVisible property false for an Excel named range with Aspose.Cells | Programmatically conceal a named range in a .NET workbook using Aspose.Cells | How to prevent a named range from appearing in Excel's Name Manager via code | C# example to hide MyNamedRange using Aspose.Cells API
// Tags: Aspose.Cells named range visibility control | Name.IsVisible false Aspose.Cells | Excel named range hide .NET API | Workbook.Worksheets.Names hide range | Aspose.Cells conceal Name Manager entry

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads (or creates) an Excel workbook, retrieves the named range "MyNamedRange", sets its IsVisible property to false to hide it from the Name Manager, saves the workbook, and includes handling for Aspose.Cells-specific and general exceptions.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                Workbook workbook;

                if (File.Exists(inputPath))
                {
                    // Load the workbook. If the file is password‑protected, set the password in LoadOptions.
                    var loadOptions = new LoadOptions();
                    // loadOptions.Password = "yourPassword"; // Uncomment and set if needed
                    workbook = new Workbook(inputPath, loadOptions);
                }
                else
                {
                    // Create a new workbook when the input file is missing.
                    workbook = new Workbook();
                }

                // Retrieve the named range.
                Name namedRange = workbook.Worksheets.Names["MyNamedRange"];
                if (namedRange != null)
                {
                    // Hide the named range so it does not appear in the Name Manager.
                    namedRange.IsVisible = false;
                }
                else
                {
                    Console.WriteLine("Named range 'MyNamedRange' not found.");
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
