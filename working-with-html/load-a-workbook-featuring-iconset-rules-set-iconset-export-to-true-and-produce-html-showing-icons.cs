// Title: Load an Excel workbook containing IconSet conditional formatting and export it to HTML with visible icons using Aspose.Cells for .NET (C#)
// Common Searches: how to export Excel IconSet conditional formatting to HTML with Aspose.Cells C# | Aspose.Cells HtmlSaveOptions ExportConditionalFormatting example for icon sets | C# code to generate HTML that shows traffic light icons from an Excel file | save workbook as HTML while preserving conditional formatting icons using Aspose.Cells | load existing workbook with IconSet rules and convert to web page in .NET
// Tags: iconset conditionalformatting html export | aspocells htmlsaveoptions exportconditionalformatting | c# load workbook with iconset | excel to html conversion preserving icons | conditional formatting icons in html output

using System;
using System.IO;
using Aspose.Cells;

namespace IconSetHtmlExportDemo
{
    // The sample checks for an Excel file, creates one with a ThreeTrafficLights1 IconSet rule if missing, loads the workbook, configures HtmlSaveOptions to keep conditional formatting, and saves the sheet as HTML so the icons appear in the generated web page.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for the input workbook and the output HTML file
            string inputFile = "IconSetDemo.xlsx";
            string outputHtml = "IconSetOutput.html";

            try
            {
                // -----------------------------------------------------------------
                // 1. Ensure a workbook with an IconSet conditional formatting exists.
                // -----------------------------------------------------------------
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file \"{inputFile}\" not found. Creating a sample workbook...");

                    // Create a new workbook and get the first worksheet
                    Workbook sampleWb = new Workbook();
                    Worksheet sheet = sampleWb.Worksheets[0];
                    Cells cells = sheet.Cells;

                    // Populate some numeric data (A1:A5)
                    for (int i = 0; i < 5; i++)
                    {
                        cells[i, 0].PutValue(10 * (i + 1)); // 10, 20, 30, 40, 50
                    }

                    // Add an IconSet conditional formatting rule to the range A1:A5
                    int cfIndex = sheet.ConditionalFormattings.Add();
                    ConditionalFormatting cf = sheet.ConditionalFormattings[cfIndex];
                    cf.AddArea(new CellArea
                    {
                        StartRow = 0,
                        EndRow = 4,
                        StartColumn = 0,
                        EndColumn = 0
                    });

                    // Create the IconSet condition
                    var condition = cf.AddCondition(ConditionType.IconSet, string.Empty, string.Empty);
                    if (condition is IconSetCondition iconSetCondition)
                    {
                        // Choose a built‑in icon set (Three Traffic Lights)
                        iconSetCondition.IconSet = IconSetType.ThreeTrafficLights1;

                        // Show both the value and the icon
                        iconSetCondition.ShowIconOnly = false;
                    }

                    // Save the sample workbook
                    sampleWb.Save(inputFile);
                    Console.WriteLine($"Sample workbook saved as \"{inputFile}\".");
                }

                // -----------------------------------------------------------------
                // 2. Load the workbook that contains the IconSet rule.
                // -----------------------------------------------------------------
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: The workbook \"{inputFile}\" does not exist.");
                    return;
                }

                Workbook workbook = new Workbook(inputFile);
                Console.WriteLine($"Workbook \"{inputFile}\" loaded successfully.");

                // -----------------------------------------------------------------
                // 3. Configure HTML save options to export conditional formatting (including IconSet).
                // -----------------------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export only the active worksheet (optional, makes the output smaller)
                    ExportActiveWorksheetOnly = true,

                    // Keep conditional formatting (including IconSet, data bars, color scales, etc.)
                    ExportConditionalFormatting = true
                };

                // -----------------------------------------------------------------
                // 4. Save the workbook as HTML.
                // -----------------------------------------------------------------
                workbook.Save(outputHtml, htmlOptions);
                Console.WriteLine($"HTML file with IconSet icons saved as \"{outputHtml}\".");

                // -----------------------------------------------------------------
                // 5. Inform the user where to find the result.
                // -----------------------------------------------------------------
                Console.WriteLine("Open the generated HTML file in a web browser to see the icons.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
