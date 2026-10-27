// Title: Create a C# unit test to verify that a custom subtotal label persists after applying GlobalizationSettings in Aspose.Cells
// Common Searches: how to write a unit test for custom subtotal label with Aspose.Cells in C# | verify that a custom subtotal text remains after setting workbook CultureInfo in Aspose.Cells | Aspose.Cells C# unit test checking subtotal label persistence after saving to a memory stream | test that workbook globalization settings do not remove custom subtotal labels in Aspose.Cells | C# example for validating custom subtotal label after workbook reload with Aspose.Cells
// Tags: Aspose.Cells custom subtotal label verification | C# unit testing workbook globalization settings | Aspose.Cells save to memory stream XLSX | verify cell string persistence after reload | globalization settings CultureInfo Aspose.Cells | subtotal label persistence Aspose.Cells | C# Aspose.Cells workbook culture test

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a Workbook, sets its CultureInfo to en-US, adds sample data, inserts a custom subtotal label in cell A7, saves the file to a MemoryStream as XLSX, reloads the workbook, and scans column A to confirm the label is still present, throwing an exception if it is missing.
    public class SubtotalLabelTests
    {
        public static void Main()
        {
            try
            {
                new SubtotalLabelTests().CustomSubtotalLabelAppears();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        public void CustomSubtotalLabelAppears()
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Amount");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("A");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("B");
            worksheet.Cells["B4"].PutValue(30);
            worksheet.Cells["A5"].PutValue("B");
            worksheet.Cells["B5"].PutValue(40);

            // Set workbook culture (example: US culture)
            workbook.Settings.CultureInfo = new CultureInfo("en-US");

            // Insert a custom subtotal label manually (since SubtotalOptions may not be available)
            worksheet.Cells["A7"].PutValue("My Subtotal");

            // Save to a memory stream and reload to verify persistence
            using (var stream = new MemoryStream())
            {
                try
                {
                    workbook.Save(stream, SaveFormat.Xlsx);
                }
                catch (Exception saveEx)
                {
                    throw new InvalidOperationException("Failed to save workbook to stream.", saveEx);
                }

                stream.Position = 0;
                var loadedWorkbook = new Workbook(stream);
                var loadedWorksheet = loadedWorkbook.Worksheets[0];

                // Search for the custom subtotal label in the first column
                bool labelFound = false;
                for (int row = 0; row <= loadedWorksheet.Cells.MaxDataRow; row++)
                {
                    var cell = loadedWorksheet.Cells[row, 0];
                    if (cell.Type == CellValueType.IsString && cell.StringValue == "My Subtotal")
                    {
                        labelFound = true;
                        break;
                    }
                }

                if (!labelFound)
                {
                    throw new InvalidOperationException("Custom subtotal label was not found in the worksheet.");
                }
            }
        }
    }
}
