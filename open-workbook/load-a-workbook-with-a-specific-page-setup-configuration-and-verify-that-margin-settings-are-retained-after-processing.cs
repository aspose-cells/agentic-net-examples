// Title: C# Example: Set and Verify Worksheet Page Margins with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, assign left, right, top, and bottom margins in centimeters via the PageSetup object, save the file, reload it, and programmatically confirm that the margin values are retained.
// Keywords: Aspose.Cells | C# page margins | Excel margin persistence | Worksheet PageSetup margins | load workbook verify margins | Aspose.Cells .NET margin example | set margins centimeters | Excel print margins C#
// Common Searches: Aspose.Cells set worksheet margins C# | How to keep Excel page margins after saving with Aspose.Cells | Verify Excel margins programmatically .NET | C# code to compare original and loaded page margins | Aspose.Cells margin retention test
// Developer Intent: Confirm that page margin settings applied to a worksheet stay unchanged after the workbook is saved and reloaded using Aspose.Cells for .NET.
// Use Cases: Generate a report template with precise print margins, save it, and later validate margins before distribution. | Automate quality checks across a batch of generated Excel files to ensure consistent layout settings. | Add margin verification to a CI/CD pipeline to catch unintended changes in page setup during development.
// AI Prompts: Write C# code with Aspose.Cells that sets left, right, top, and bottom margins in centimeters and verifies they are unchanged after reloading the workbook. | Provide a reusable method that compares expected margin values with those read from a loaded worksheet and returns a boolean result. | Explain how Aspose.Cells stores page margin values, the default unit, and how to convert between centimeters and points.

using System;
using Aspose.Cells;

namespace AsposeCellsMarginVerification
{
    // Demonstrates how to create a workbook, assign left, right, top, and bottom margins in centimeters via the PageSetup object, save the file, reload it, and programmatically confirm that the margin values are retained.
    class Program
    {
        static void Main()
        {
            // Define the path for the temporary workbook
            string filePath = "MarginTest.xlsx";

            // -------------------- Create and configure workbook --------------------
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set page margin values (centimeters)
            double leftMargin = 1.0;
            double rightMargin = 1.5;
            double topMargin = 2.0;
            double bottomMargin = 0.5;

            sheet.PageSetup.LeftMargin = leftMargin;
            sheet.PageSetup.RightMargin = rightMargin;
            sheet.PageSetup.TopMargin = topMargin;
            sheet.PageSetup.BottomMargin = bottomMargin;

            // Save the workbook to disk
            workbook.Save(filePath, SaveFormat.Xlsx);

            // -------------------- Load workbook and verify margins --------------------
            // Load the workbook from the saved file
            Workbook loadedWorkbook = new Workbook(filePath);

            // Access the first worksheet of the loaded workbook
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            PageSetup loadedSetup = loadedSheet.PageSetup;

            // Retrieve margin values from the loaded workbook
            double loadedLeft = loadedSetup.LeftMargin;
            double loadedRight = loadedSetup.RightMargin;
            double loadedTop = loadedSetup.TopMargin;
            double loadedBottom = loadedSetup.BottomMargin;

            // Verify that the margins match the original values
            bool marginsMatch = Math.Abs(loadedLeft - leftMargin) < 0.0001 &&
                                Math.Abs(loadedRight - rightMargin) < 0.0001 &&
                                Math.Abs(loadedTop - topMargin) < 0.0001 &&
                                Math.Abs(loadedBottom - bottomMargin) < 0.0001;

            // Output verification result
            Console.WriteLine("Margin verification result: " + (marginsMatch ? "Success" : "Failure"));
            Console.WriteLine($"LeftMargin: Expected={leftMargin}, Loaded={loadedLeft}");
            Console.WriteLine($"RightMargin: Expected={rightMargin}, Loaded={loadedRight}");
            Console.WriteLine($"TopMargin: Expected={topMargin}, Loaded={loadedTop}");
            Console.WriteLine($"BottomMargin: Expected={bottomMargin}, Loaded={loadedBottom}");
        }
    }
}
