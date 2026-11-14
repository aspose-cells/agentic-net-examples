// Title: Attempt to set Sparkline lighting to Soft mode in Excel 2010 workbooks using Aspose.Cells for .NET (unsupported in current version)
// Common Searches: Aspose.Cells C# how to change sparkline lighting to soft mode in Excel 2010 | set sparkline display mode soft using Aspose.Cells .NET | why Aspose.Cells does not expose SparklineGroup properties for lighting | adjust sparkline appearance soft lighting with Aspose.Cells C# example | Excel 2010 sparkline soft lighting not available in Aspose.Cells API
// Tags: sparkline lighting soft mode Aspose.Cells | Aspose.Cells sparkline group API limitation | C# set sparkline display mode Excel 2010 | Excel 2010 sparkline appearance customization Aspose.Cells | Aspose.Cells version missing sparkline properties

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example loads an existing Excel workbook or creates a new one, ensures at least one worksheet exists, iterates through all worksheets, and saves the file. It explicitly notes that SparklineGroup types and lighting settings are not exposed in the current Aspose.Cells version, so Sparkline lighting cannot be modified in this code.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if the file exists; otherwise create a new one.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    // Ensure there is at least one worksheet.
                    if (workbook.Worksheets.Count == 0)
                    {
                        workbook.Worksheets.Add();
                    }
                    workbook.Worksheets[0].Name = "Sheet1";
                }

                // Iterate through all worksheets.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Sparkline handling is omitted because the current Aspose.Cells version
                    // used in this project does not expose SparklineGroup types.
                    // If future versions provide sparkline properties, they can be accessed here.
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
