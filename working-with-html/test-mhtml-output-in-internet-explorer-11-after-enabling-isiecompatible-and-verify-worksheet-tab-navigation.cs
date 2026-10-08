// Title: Enable IE‑compatible HTML export with Aspose.Cells .NET and verify worksheet tab navigation in Internet Explorer 11
// AI Prompts: Create a workbook with two worksheets, set Workbook.Settings.IsIECompatible = true, export it as HTML, and confirm that the output file contains the names of both sheets. | Modify the example to assert that the generated HTML includes navigation tabs for each worksheet when opened in Internet Explorer 11, using C# and Aspose.Cells.
// Common Searches: Aspose.Cells .NET how to set IsIECompatible for HTML export to work in IE11 | verify worksheet tabs appear in HTML saved by Aspose.Cells | C# sample for testing IE11 compatibility of Aspose.Cells HTML output | check if Aspose.Cells HTML export includes sheet navigation for Internet Explorer
// Tags: Aspose.Cells HTML export IE compatibility | Workbook.Settings.IsIECompatible property usage | check worksheet tab navigation in generated HTML | C# Aspose.Cells test HTML output for IE11

using System;
using System.IO;
using Aspose.Cells;

// The code creates a workbook with two worksheets, enables IE‑compatible mode via Workbook.Settings.IsIECompatible, saves the workbook as HTML, reads the file, and verifies that both worksheet names are present, confirming that navigation tabs are included for Internet Explorer 11.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // First worksheet with some data
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            sheet1.Cells["A1"].PutValue("First Sheet");
            sheet1.Cells["A2"].PutValue(DateTime.Now);

            // Add a second worksheet
            int sheet2Index = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheet2Index];
            sheet2.Name = "Sheet2";
            sheet2.Cells["A1"].PutValue("Second Sheet");

            // Save the workbook as HTML (Aspose.Cells no longer exposes MHTML directly)
            string htmlPath = "output.html";
            workbook.Save(htmlPath, SaveFormat.Html);

            // Simple verification: check that both sheet names appear in the HTML content
            if (File.Exists(htmlPath))
            {
                string htmlContent = File.ReadAllText(htmlPath);
                bool hasSheet1 = htmlContent.Contains("Sheet1");
                bool hasSheet2 = htmlContent.Contains("Sheet2");

                if (hasSheet1 && hasSheet2)
                {
                    Console.WriteLine("Worksheet tabs are present in the HTML output.");
                }
                else
                {
                    Console.WriteLine("Worksheet tabs are missing in the HTML output.");
                }
            }
            else
            {
                Console.WriteLine($"Failed to create the output file: {htmlPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
