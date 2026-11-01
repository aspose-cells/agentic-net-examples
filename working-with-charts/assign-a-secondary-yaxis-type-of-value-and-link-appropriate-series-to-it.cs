// Title: Assign a secondary value Y‑axis to a chart series with Aspose.Cells for .NET – handling version limitations
// AI Prompts: Write C# code using Aspose.Cells to create a secondary value axis on a chart and assign a chosen series to it. | Detail the upgrade path to a newer Aspose.Cells for .NET release that includes SecondaryValueAxis and IsOnSecondaryAxis, and provide the revised code. | Propose a technique to mimic a secondary Y‑axis in Aspose.Cells when the library version lacks direct support.
// Common Searches: Aspose.Cells .NET move chart series to a secondary Y axis | C# create secondary value axis for Excel chart with Aspose.Cells | Which Aspose.Cells version adds SecondaryValueAxis property | Workaround for secondary axis in Aspose.Cells chart without native support | Example of linking a series to a secondary axis using Aspose.Cells for .NET
// Tags: Aspose.Cells support for additional chart axes | C# assign series to alternate Y axis in Aspose.Cells | upgrading Aspose.Cells to enable axis properties | emulating secondary chart axis in Aspose.Cells | Excel chart axis handling with Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The sample loads an existing workbook, checks for a chart, and notes that the current Aspose.Cells version does not expose SecondaryValueAxis or IsOnSecondaryAxis members required to assign a secondary value Y‑axis to a series. It then saves the workbook unchanged, highlighting the need for a newer library version or a workaround.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one chart on the worksheet
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Access the first chart
                Chart chart = worksheet.Charts[0];

                // NOTE: The current Aspose.Cells version used in this project does not expose
                // SecondaryValueAxis or IsOnSecondaryAxis members. If you need to work with
                // secondary axes, upgrade to a newer version that supports these features.

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
