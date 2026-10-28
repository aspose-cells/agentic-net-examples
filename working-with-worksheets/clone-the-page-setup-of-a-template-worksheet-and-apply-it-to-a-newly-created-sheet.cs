// Title: How to clone page‑setup settings from a template worksheet to a new worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy page setup from one worksheet to another | duplicate worksheet print margins and paper size with Aspose.Cells .NET | clone print area and title rows in a new sheet using Aspose.Cells API | transfer page setup properties between worksheets in C# Aspose.Cells | copy worksheet page setup settings programmatically with Aspose.Cells for .NET
// Tags: Aspose.Cells worksheet page setup cloning | C# copy print margins and paper size | Aspose.Cells transfer print area between sheets | clone worksheet print titles .NET | apply template page setup programmatically

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a template workbook (or creates one if missing), accesses its first worksheet, adds a new worksheet, and copies all supported PageSetup properties—including paper size, orientation, margins, print area, and title rows/columns—from the template sheet to the new sheet before saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "template.xlsx";
                const string outputPath = "output.xlsx";

                // Load existing template or create a new workbook if the file is missing
                Workbook workbook;
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    workbook = new Workbook();
                    workbook.Worksheets[0].Name = "Template";
                }

                // Reference the template worksheet (first sheet)
                Worksheet templateSheet = workbook.Worksheets[0];

                // Add a new worksheet that will receive the cloned page setup
                Worksheet newSheet = workbook.Worksheets.Add("ClonedPageSetup");

                // Clone supported page‑setup properties
                newSheet.PageSetup.PaperSize = templateSheet.PageSetup.PaperSize;
                newSheet.PageSetup.Orientation = templateSheet.PageSetup.Orientation;
                newSheet.PageSetup.FitToPagesWide = templateSheet.PageSetup.FitToPagesWide;
                newSheet.PageSetup.FitToPagesTall = templateSheet.PageSetup.FitToPagesTall;
                newSheet.PageSetup.PrintArea = templateSheet.PageSetup.PrintArea;
                newSheet.PageSetup.PrintTitleRows = templateSheet.PageSetup.PrintTitleRows;
                newSheet.PageSetup.PrintTitleColumns = templateSheet.PageSetup.PrintTitleColumns;
                newSheet.PageSetup.LeftMargin = templateSheet.PageSetup.LeftMargin;
                newSheet.PageSetup.RightMargin = templateSheet.PageSetup.RightMargin;
                newSheet.PageSetup.TopMargin = templateSheet.PageSetup.TopMargin;
                newSheet.PageSetup.BottomMargin = templateSheet.PageSetup.BottomMargin;
                newSheet.PageSetup.HeaderMargin = templateSheet.PageSetup.HeaderMargin;
                newSheet.PageSetup.FooterMargin = templateSheet.PageSetup.FooterMargin;
                newSheet.PageSetup.CenterHorizontally = templateSheet.PageSetup.CenterHorizontally;
                newSheet.PageSetup.CenterVertically = templateSheet.PageSetup.CenterVertically;
                newSheet.PageSetup.PrintGridlines = templateSheet.PageSetup.PrintGridlines;
                newSheet.PageSetup.PrintHeadings = templateSheet.PageSetup.PrintHeadings;
                newSheet.PageSetup.BlackAndWhite = templateSheet.PageSetup.BlackAndWhite;

                // Save the workbook with the cloned page setup
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
