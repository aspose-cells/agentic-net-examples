// Title: Save an existing workbook that contains a chart as an XLSX file to a specific output folder using Aspose.Cells for .NET
// AI Prompts: Load a workbook that includes a chart and export it to a given directory as an XLSX file with Aspose.Cells in C#. | Create the destination folder if it does not exist, then save the opened workbook using SaveFormat.Xlsx to the target path. | Write C# code that opens a source Excel file, ensures the output directory is present, and writes the workbook containing charts to that location with Aspose.Cells.
// Common Searches: Aspose.Cells C# save workbook with chart to custom folder | How to export an Excel file that contains charts to a specific directory using Aspose.Cells | Create output directory before saving workbook as XLSX in .NET | Load existing Excel workbook and save as XLSX with Aspose.Cells SaveFormat.Xlsx
// Tags: save workbook with chart Aspose.Cells | export Excel to custom directory C# | Aspose.Cells SaveFormat.Xlsx usage | ensure output folder exists C# | load existing workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads a workbook that already has a chart, creates the target output folder if needed, and saves the workbook as an XLSX file to the specified path using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Specify the output directory where the XLSX file will be saved
        string outputDirectory = @"C:\Output";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDirectory);

        // Define the full path for the saved workbook
        string outputPath = Path.Combine(outputDirectory, "ChartWorkbook.xlsx");

        // Load the workbook that already contains the chart
        // (Replace the source path with the actual location of your workbook)
        string sourceWorkbookPath = @"C:\Input\SourceWorkbook.xlsx";
        Workbook workbook = new Workbook(sourceWorkbookPath);

        // Save the workbook as an XLSX file to the specified output directory
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
