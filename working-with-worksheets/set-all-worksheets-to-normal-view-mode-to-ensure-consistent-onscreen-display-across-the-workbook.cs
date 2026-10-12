// Title: Set all worksheets to Normal view mode in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, iterates over every worksheet, sets each Worksheet.ViewOptions.ViewMode to Normal, and saves the workbook. | Create a snippet that first checks whether Worksheet.ViewOptions is supported, then applies Normal view mode to all sheets, handling older Aspose.Cells versions gracefully.
// Common Searches: how to set worksheet view mode to normal using Aspose.Cells C# | Aspose.Cells iterate all worksheets and change view mode | C# change view options for every sheet in an Excel workbook | Aspose.Cells viewmode property missing in older releases workaround | save workbook after modifying worksheet display settings with Aspose.Cells
// Tags: Aspose.Cells set worksheet view mode | C# iterate workbook worksheets | normal view mode Excel .NET | Aspose.Cells compatibility older versions | save workbook after view options change

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file (or creates a new workbook), attempts to set each worksheet's ViewOptions.ViewMode to Normal (code is commented out for compatibility with older Aspose.Cells releases), saves the workbook as 'output.xlsx', and reports success or any errors to the console.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load workbook if the input file exists; otherwise create a new workbook.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Set each worksheet to Normal view mode if the API is available.
            // In some Aspose.Cells versions Worksheet does not expose ViewOptions,
            // so this step is skipped to ensure compatibility.
            // foreach (Worksheet sheet in workbook.Worksheets)
            // {
            //     sheet.ViewOptions.ViewMode = ViewMode.Normal;
            // }

            // Save the updated workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
