// Title: How to save an Excel workbook as XLSX while excluding hidden worksheets using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing .xlsx file with Aspose.Cells, sets ExportHiddenWorksheet = false on OoxmlSaveOptions, and saves a new workbook containing only visible sheets. | Explain the steps to configure OoxmlSaveOptions in Aspose.Cells so that hidden worksheets are not written to the output file. | Create a sample program that validates the source file, applies the ExportHiddenWorksheet setting, and writes the result to a specified output path.
// Common Searches: Aspose.Cells C# export only visible worksheets to XLSX | How to prevent hidden sheets from being saved with Aspose.Cells | Using OoxmlSaveOptions to omit hidden worksheets in .NET | Save Excel file without hidden sheets using Aspose.Cells library | C# example of hiding hidden worksheets during workbook export
// Tags: Aspose.Cells OoxmlSaveOptions hidden sheet exclusion | C# save workbook visible sheets only | exclude hidden worksheets during XLSX export | Aspose.Cells hide hidden worksheets on save | Excel export without hidden sheets .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a workbook, verifies the input file, creates OoxmlSaveOptions for the XLSX format, and saves the workbook to a new file. Although the task describes setting ExportHiddenWorksheet to false to exclude hidden sheets, this property is not available in the demonstrated version of Aspose.Cells, so the code saves all worksheets.
class Program
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

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Configure save options (hidden worksheets will be included as the property is unavailable in this version)
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);

            // Save the workbook to a new file with the specified options
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
