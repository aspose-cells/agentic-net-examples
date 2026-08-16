// Title: Verify Excel PageSetup Margins After Load/Save with Aspose.Cells for .NET
// Description: Demonstrates how to create an Excel workbook, set left, right, top, and bottom margins in centimeters, save it, reload it with Aspose.Cells, read the PageSetup margin values, confirm they match the expected measurements, add a confirmation cell, and save a processed copy while preserving the original margin configuration.
// Keywords: Aspose.Cells margin verification | C# Excel PageSetup margins | load workbook Aspose.Cells .NET | preserve Excel margins after save | PageSetup LeftMargin RightMargin TopMargin BottomMargin | Excel margin unit centimeters | Aspose.Cells margin unit conversion
// Common Searches: how to read page margins from Excel using Aspose.Cells | verify Excel margin values after saving with Aspose.Cells | Aspose.Cells get left and right margin in cm | C# check Excel PageSetup margins after load | preserve page setup settings Aspose.Cells .NET
// Developer Intent: Ensure that margin settings defined in an Excel workbook remain unchanged after loading, processing, and re‑saving with Aspose.Cells.
// Use Cases: Set custom page margins for a workbook before distribution and confirm they are retained. | Load an existing Excel file, read its PageSetup margins, and validate them against design specifications. | Perform additional worksheet operations after margin verification while keeping the original margin configuration intact.
// AI Prompts: Generate C# code using Aspose.Cells to set page margins in centimeters, save the workbook, reload it, and assert that LeftMargin, RightMargin, TopMargin, and BottomMargin are unchanged. | Write a .NET unit test that opens an Excel file with Aspose.Cells and verifies that all PageSetup margin values match expected numbers. | Explain how Aspose.Cells stores PageSetup margin values, the default units, and how unit conversion works when reading or writing margins.

using System;
using Aspose.Cells;

namespace PageSetupMarginVerification
{
    // Demonstrates how to create an Excel workbook, set left, right, top, and bottom margins in centimeters, save it, reload it with Aspose.Cells, read the PageSetup margin values, confirm they match the expected measurements, add a confirmation cell, and save a processed copy while preserving the original margin configuration.
    class Program
    {
        static void Main()
        {
            // Define file paths
            string originalPath = "OriginalMargins.xlsx";
            string processedPath = "ProcessedMargins.xlsx";

            // ---------- Create a new workbook and set margin values ----------
            Workbook workbook = new Workbook(); // create
            Worksheet sheet = workbook.Worksheets[0];

            // Set specific margin values (centimeters)
            sheet.PageSetup.LeftMargin = 1.0;    // 1 cm
            sheet.PageSetup.RightMargin = 1.5;   // 1.5 cm
            sheet.PageSetup.TopMargin = 2.0;     // 2 cm
            sheet.PageSetup.BottomMargin = 2.5;  // 2.5 cm

            // Save the workbook with the configured margins
            workbook.Save(originalPath, SaveFormat.Xlsx); // save

            // ---------- Load the workbook and verify margin settings ----------
            Workbook loadedWorkbook = new Workbook(originalPath); // load
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            PageSetup ps = loadedSheet.PageSetup;

            // Retrieve margin values
            double left = ps.LeftMargin;
            double right = ps.RightMargin;
            double top = ps.TopMargin;
            double bottom = ps.BottomMargin;

            // Output the retrieved margin values
            Console.WriteLine($"Loaded Margins (cm): Left={left}, Right={right}, Top={top}, Bottom={bottom}");

            // Simple verification: compare with expected values
            bool marginsMatch = Math.Abs(left - 1.0) < 0.0001 &&
                                Math.Abs(right - 1.5) < 0.0001 &&
                                Math.Abs(top - 2.0) < 0.0001 &&
                                Math.Abs(bottom - 2.5) < 0.0001;

            Console.WriteLine("Margin verification " + (marginsMatch ? "succeeded." : "failed."));

            // Optionally, perform additional processing (e.g., add data) and save again
            loadedSheet.Cells["A1"].PutValue("Margin verification completed.");
            loadedWorkbook.Save(processedPath, SaveFormat.Xlsx); // save processed workbook
        }
    }
}
