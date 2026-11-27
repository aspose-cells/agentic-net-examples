// Title: Create an internal worksheet hyperlink that jumps to Sheet1!A1 from Sheet2 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add internal hyperlink that opens a specific cell on another sheet | How to use '#Sheet1!A1' syntax for internal links in an Aspose.Cells workbook | Set a screen tip for an internal hyperlink in an Excel file with Aspose.Cells C# | Create hyperlink between worksheets and navigate to a target cell using Aspose.Cells | Verify internal hyperlink address after saving a workbook with Aspose.Cells .NET
// Tags: Aspose.Cells internal worksheet hyperlink | C# add hyperlink '#SheetName!CellReference' Aspose.Cells | Aspose.Cells set hyperlink screen tip | Aspose.Cells navigate to cell via hyperlink | Aspose.Cells save workbook with internal link

using System;
using System.IO;
using Aspose.Cells;

namespace HyperlinkInternalDemo
{
    // The example creates a new workbook with two worksheets, writes sample data, and adds an internal hyperlink on Sheet2!B2 that points to Sheet1!A1 using the '#Sheet1!A1' address format. It also sets a screen tip for the link, saves the workbook as an XLSX file, reloads it, and prints the hyperlink's address, display text, and screen tip to verify correctness.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook with a default worksheet named "Sheet1"
                Workbook workbook = new Workbook();

                // Add a second worksheet and set its name to "Sheet2"
                int sheet2Index = workbook.Worksheets.Add();
                Worksheet sheet2 = workbook.Worksheets[sheet2Index];
                sheet2.Name = "Sheet2";

                // Get reference to the first worksheet ("Sheet1")
                Worksheet sheet1 = workbook.Worksheets[0];

                // Populate some data so we can see the target cell
                sheet1.Cells["A1"].PutValue("Target Cell on Sheet1");
                sheet2.Cells["B2"].PutValue("Click to go to Sheet1!A1");

                // Add an internal hyperlink on Sheet2!B2 that points to Sheet1!A1
                // The address uses the '#SheetName!CellReference' syntax.
                // Parameters: cell name, first row, first column (zero‑based), hyperlink address
                sheet2.Hyperlinks.Add("B2", 1, 1, "#Sheet1!A1");

                // Optionally, set a screen tip for the hyperlink
                Hyperlink link = sheet2.Hyperlinks[sheet2.Hyperlinks.Count - 1];
                link.ScreenTip = "Navigate to Sheet1 cell A1";

                // Save the workbook
                string outputPath = "HyperlinkInternalDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");

                // Verify the hyperlink was stored correctly
                if (File.Exists(outputPath))
                {
                    Workbook loadedWorkbook = new Workbook(outputPath);
                    Worksheet loadedSheet2 = loadedWorkbook.Worksheets["Sheet2"];
                    Hyperlink loadedLink = loadedSheet2.Hyperlinks[0];

                    Console.WriteLine("Verified hyperlink details:");
                    Console.WriteLine($"  Address   : {loadedLink.Address}");
                    // The display text is in cell B2
                    Console.WriteLine($"  TextToDisplay (cell B2) : {loadedSheet2.Cells["B2"].StringValue}");
                    Console.WriteLine($"  ScreenTip : {loadedLink.ScreenTip}");
                }
                else
                {
                    Console.WriteLine($"Error: The file '{outputPath}' was not found for verification.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the hyperlink:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
