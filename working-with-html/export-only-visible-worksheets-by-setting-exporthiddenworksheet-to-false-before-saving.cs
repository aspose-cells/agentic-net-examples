// Title: Export only visible worksheets to a new XLSX file with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook using Aspose.Cells, removes every worksheet whose IsVisible property is false, and saves the remaining sheets to a new .xlsx file. | Demonstrate how to prevent hidden worksheets from being saved by filtering them out before calling Workbook.Save in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells .NET export visible worksheets only | How to save an Excel file without hidden sheets using C# | Remove hidden worksheets from workbook before saving with Aspose.Cells | C# code to filter out invisible worksheets in Aspose.Cells | Save workbook with only visible sheets Aspose.Cells SaveFormat.Xlsx
// Tags: export visible worksheets Aspose.Cells | remove hidden worksheets C# | save workbook without hidden sheets .NET | filter worksheets by visibility Aspose.Cells | exclude hidden sheets during save Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file with Aspose.Cells, identifies worksheets where IsVisible is false, removes those hidden worksheets, and then saves the workbook so that only the visible sheets are retained in the new XLSX output.
class ExportVisibleWorksheets
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Collect indexes of hidden worksheets
            List<int> hiddenIndexes = new List<int>();
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                if (workbook.Worksheets[i].IsVisible == false)
                {
                    hiddenIndexes.Add(i);
                }
            }

            // Remove hidden worksheets starting from the last index to avoid shifting
            for (int i = hiddenIndexes.Count - 1; i >= 0; i--)
            {
                workbook.Worksheets.RemoveAt(hiddenIndexes[i]);
            }

            // Save the workbook (only visible worksheets remain)
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
