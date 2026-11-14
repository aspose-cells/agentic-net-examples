// Title: Hide the third worksheet, note TabBarWidth limitation, and save the workbook using Aspose.Cells for .NET
// Common Searches: how to hide a specific worksheet in an Aspose.Cells workbook using C# | does Aspose.Cells provide an API to set Excel tab bar width programmatically | save a workbook after changing worksheet visibility with Aspose.Cells .NET | Aspose.Cells C# example to hide the third sheet and export to XLSX | workaround for missing TabBarWidth property in Aspose.Cells
// Tags: worksheet visibility control Aspose.Cells | tab bar width unsupported Aspose.Cells | export workbook to XLSX Aspose.Cells | Excel UI customization limitations Aspose.Cells | C# hide worksheet using Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, hides the third worksheet (index 2) by setting its IsVisible property to false, documents that the TabBarWidth property is not available in the current Aspose.Cells API, and then saves the workbook as Result.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (contains three default worksheets)
                Workbook workbook = new Workbook();

                // Hide the third worksheet (index 2, zero‑based)
                Worksheet thirdSheet = workbook.Worksheets[2];
                thirdSheet.IsVisible = false; // or thirdSheet.Visibility = WorksheetVisibility.Hidden;

                // The TabBarWidth property is not available in the current Aspose.Cells API.
                // If required, adjust related settings using supported properties.

                // Save the workbook to a file
                string outputPath = "Result.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
