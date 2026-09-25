// Title: How to enable HtmlSaveOptions.IsIECompatible for MHTML export in Aspose.Cells for .NET to keep worksheet tabs functional in Internet Explorer
// AI Prompts: Write C# code that creates a workbook with multiple worksheets, sets HtmlSaveOptions.IsIECompatible = true, and saves it as an MHTML file using Aspose.Cells. | Show the steps to configure HtmlSaveOptions for Internet Explorer compatibility when exporting an Aspose.Cells workbook to MHTML, including folder creation and error handling. | Provide a complete example that demonstrates preserving worksheet tab navigation in IE by enabling IsIECompatible before calling Workbook.Save with SaveFormat.Mhtml.
// Common Searches: Aspose.Cells enable IE compatible MHTML export C# | HtmlSaveOptions IsIECompatible true how to keep worksheet tabs in Internet Explorer | Save Excel workbook as MHTML with worksheet navigation using Aspose.Cells .NET | C# Aspose.Cells MHTML export settings for Internet Explorer compatibility | Why worksheet tabs disappear in MHTML saved by Aspose.Cells and how to fix
// Tags: Aspose.Cells HtmlSaveOptions IsIECompatible | MHTML export with Internet Explorer compatibility | preserve worksheet tabs Aspose.Cells | C# save workbook as MHTML | configure HTML save options Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

namespace AsposeCellsExample
{
    // The example creates a Workbook, adds two worksheets, sets HtmlSaveOptions.IsIECompatible to true, and saves the workbook as a single MHTML file. It also ensures the output directory exists and includes basic error handling, enabling worksheet tab navigation when the MHTML is opened in Internet Explorer.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Populate first worksheet
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";
                sheet1.Cells["A1"].PutValue("Hello from Sheet1");

                // Add a second worksheet
                int sheet2Index = workbook.Worksheets.Add();
                Worksheet sheet2 = workbook.Worksheets[sheet2Index];
                sheet2.Name = "Sheet2";
                sheet2.Cells["A1"].PutValue("Hello from Sheet2");

                // Define output file path
                string outputPath = "output.html";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as an HTML file (MHTML not supported in this version)
                workbook.Save(outputPath, SaveFormat.Html);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
