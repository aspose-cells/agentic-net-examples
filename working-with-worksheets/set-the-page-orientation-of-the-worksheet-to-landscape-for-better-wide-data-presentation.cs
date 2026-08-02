// Title: Set Worksheet Page Orientation to Landscape Using Aspose.Cells for .NET
// Description: Shows how to create a workbook, access its first worksheet, set PageSetup.Orientation to Landscape for wide‑data printing, and save the result as an XLSX file.
// Keywords: Aspose.Cells landscape orientation | C# set worksheet page orientation | Aspose.Cells PageSetup Landscape | print wide Excel sheet Aspose.Cells | change Excel print orientation .NET
// Common Searches: Aspose.Cells set page orientation to landscape C# | How to print Excel worksheet in landscape with Aspose.Cells | C# Aspose.Cells change worksheet print layout | Landscape page setup Aspose.Cells example | Set orientation for all worksheets Aspose.Cells
// Developer Intent: Configure a worksheet’s print layout to landscape so that wide tables or charts fit on a single printed page.
// Use Cases: Print large tables without column truncation by switching to landscape before saving. | Generate landscape‑oriented PDF reports directly from Excel worksheets. | Create printable dashboards that contain wide charts or images. | Prepare batch‑exported workbooks where each sheet must use a landscape layout.
// AI Prompts: Provide C# code to set landscape orientation for every worksheet in a workbook using Aspose.Cells. | Show how to automatically choose portrait or landscape based on column count with Aspose.Cells PageSetup. | Explain the impact of PageSetup.Orientation on PDF and image export formats in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsOrientationDemo
{
    // Shows how to create a workbook, access its first worksheet, set PageSetup.Orientation to Landscape for wide‑data printing, and save the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the page orientation to Landscape for better wide data presentation
            worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // (Optional) Add some data to visualize the effect
            worksheet.Cells["A1"].PutValue("Landscape Orientation Demo");
            for (int i = 1; i <= 10; i++)
            {
                worksheet.Cells[$"A{i + 1}"].PutValue($"Data Row {i}");
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("LandscapeOrientationDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
