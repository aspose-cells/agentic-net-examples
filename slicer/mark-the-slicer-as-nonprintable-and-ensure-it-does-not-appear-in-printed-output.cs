// Title: How to hide an Excel slicer from printed output using Aspose.Cells for .NET (C#)
// AI Prompts: Set the slicer's IsPrint property to false with Aspose.Cells in C#. | Use reflection to disable printing for a slicer when the IsPrint property is not directly accessible. | Check for the IsPrint property on a worksheet slicer and make it non‑printable before saving the workbook.
// Common Searches: Aspose.Cells C# make slicer not appear in print preview | disable printing of Excel slicer using Aspose.Cells API | set slicer IsPrint false via reflection Aspose.Cells .NET
// Tags: Aspose.Cells slicer non printable | C# Aspose.Cells hide slicer from print | Excel slicer IsPrint property reflection | Aspose.Cells modify slicer print setting | non printable slicer .NET workbook

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel file, accesses the first worksheet's slicer collection, uses reflection to set the slicer's IsPrint property to false when available, and saves the workbook so the slicer is omitted from printed output.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Retrieve the slicer collection
                var slicers = worksheet.Slicers;

                if (slicers.Count > 0)
                {
                    var slicer = slicers[0]; // Or locate by name: slicers["SlicerName"]

                    // Attempt to set the slicer as non‑printable using reflection.
                    // Some Aspose.Cells versions expose the IsPrint property; others do not.
                    var isPrintProp = slicer.GetType().GetProperty("IsPrint");
                    if (isPrintProp != null && isPrintProp.CanWrite)
                    {
                        isPrintProp.SetValue(slicer, false);
                    }
                    else
                    {
                        Console.WriteLine("Slicer.IsPrint property is unavailable in this Aspose.Cells version.");
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
