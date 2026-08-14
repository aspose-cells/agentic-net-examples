// Title: Aspose.Cells for .NET – Set Worksheet Default Column Width to 12 Characters (C#)
// Description: C# example that creates a workbook with Aspose.Cells, sets the worksheet's default column width to 12 characters via the Cells.StandardWidth property, prints the applied width, and saves the file as DefaultColumnWidth.xlsx.
// Keywords: Aspose.Cells | C# | .NET | default column width | Cells.StandardWidth | worksheet column width | Excel column width 12 characters | set standard width | Aspose.Cells example | consistent layout
// Common Searches: Aspose.Cells set default column width C# | How to change worksheet standard width in Aspose.Cells .NET | Set column width for all columns Aspose.Cells | Default column width 12 characters Excel using Aspose | C# code to set worksheet column width globally
// Developer Intent: Set the worksheet’s default column width to 12 characters using Aspose.Cells for .NET.
// Use Cases: Create a new workbook with a uniform column width before populating data. | Apply a predefined column width to an existing worksheet to match a template. | Standardize column sizing across multiple worksheets for consistent PDF or Excel output. | Prepare Excel files for printing where columns must have a fixed character width.
// AI Prompts: Generate C# code using Aspose.Cells that sets Cells.StandardWidth = 12 for every worksheet in a workbook and saves it as 'Report.xlsx'. | Provide an example that reads the current default column width, changes it to 12 characters, prints the new value to the console, and saves the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook with Aspose.Cells, sets the worksheet's default column width to 12 characters via the Cells.StandardWidth property, prints the applied width, and saves the file as DefaultColumnWidth.xlsx.
    public class SetDefaultColumnWidth
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the default column width to 12 characters (standard width)
            worksheet.Cells.StandardWidth = 12.0;

            // Optional: verify the setting
            Console.WriteLine("Standard Width set to: " + worksheet.Cells.StandardWidth);

            // Save the workbook (lifecycle: save)
            workbook.Save("DefaultColumnWidth.xlsx");
        }
    }
}
