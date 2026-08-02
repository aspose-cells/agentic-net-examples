// Title: C# – Read the subtitle of the first chart in an ODS workbook with Aspose.Cells
// Description: Loads an ODS workbook, accesses the first worksheet, verifies chart presence, retrieves the first chart, reads its SubTitle.Text property, and prints the subtitle to the console. Demonstrates ODS‑specific chart subtitle support in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | ODS | chart subtitle | read chart subtitle | SubTitle.Text | load ODS workbook | Aspose.Cells .NET | chart metadata extraction | extract chart subtitle
// Common Searches: Aspose.Cells get chart subtitle C# | read ODS chart subtitle .NET | how to retrieve chart subtitle from ODS file using Aspose | C# code to read chart subtitle in ODS workbook | Aspose.Cells chart subtitle property example
// Developer Intent: Obtain the subtitle text of the first chart in an ODS file using Aspose.Cells for .NET.
// Use Cases: Display the chart subtitle in a custom UI after loading an ODS report. | Validate that the chart subtitle follows naming conventions before publishing the workbook. | Log each chart's subtitle for audit or debugging purposes.
// AI Prompts: Write C# code that loads an ODS workbook and returns the subtitle of every chart, handling missing or empty subtitles. | Generate an example that updates a chart's subtitle in an ODS file and saves the changes with Aspose.Cells. | Explain how Aspose.Cells exposes the SubTitle.Text property for ODS charts and why this property is not available for other spreadsheet formats.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an ODS workbook, accesses the first worksheet, verifies chart presence, retrieves the first chart, reads its SubTitle.Text property, and prints the subtitle to the console. Demonstrates ODS‑specific chart subtitle support in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the ODS workbook from file
        Workbook workbook = new Workbook("input.ods");

        // Get the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Check that the worksheet contains at least one chart
        if (worksheet.Charts.Count > 0)
        {
            // Retrieve the first chart in the collection
            Chart chart = worksheet.Charts[0];

            // Read the subtitle text (available only for ODS files)
            string subtitleText = chart.SubTitle.Text;

            // Output the subtitle text
            Console.WriteLine("Chart subtitle: " + subtitleText);
        }
        else
        {
            Console.WriteLine("No charts found in the workbook.");
        }
    }
}
