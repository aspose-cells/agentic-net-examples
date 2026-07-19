// Title: Read the subtitle of the first chart in an ODS workbook using Aspose.Cells for .NET (C#)
// Description: Load an ODS spreadsheet with Aspose.Cells, access the first worksheet, verify chart presence, retrieve the first chart, and extract its SubTitle.Text property. The example prints the subtitle to the console and demonstrates handling when no charts exist.
// Keywords: Aspose.Cells ODS chart subtitle | C# read chart subtitle | Aspose.Cells get chart subtitle .NET | load ODS workbook Aspose | chart SubTitle.Text C#
// Common Searches: how to read chart subtitle from ODS file using Aspose.Cells | Aspose.Cells .NET retrieve first chart subtitle | C# extract subtitle text of chart in ODS workbook | Aspose.Cells chart subtitle property example
// Developer Intent: Extract the subtitle text of the first chart in an ODS file.
// Use Cases: Show the chart subtitle in a console tool for quick verification. | Compare the subtitle against an expected value in automated tests. | Collect chart subtitles while processing a batch of ODS files for metadata reporting.
// AI Prompts: Generate C# code that opens an ODS workbook with Aspose.Cells and prints the subtitle of the first chart. | Create a robust example that checks for charts and handles missing subtitles gracefully. | Write a reusable method that returns the subtitle text given a worksheet index and chart index in an ODS workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load an ODS spreadsheet with Aspose.Cells, access the first worksheet, verify chart presence, retrieve the first chart, and extract its SubTitle.Text property. The example prints the subtitle to the console and demonstrates handling when no charts exist.
class Program
{
    static void Main()
    {
        // Load the ODS workbook (uses Aspose.Cells load rule)
        Workbook workbook = new Workbook("input.ods");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Check that the worksheet contains at least one chart
        if (worksheet.Charts.Count > 0)
        {
            // Retrieve the first chart
            Chart chart = worksheet.Charts[0];

            // Read the subtitle text (available only for ODS format)
            string subtitleText = chart.SubTitle.Text;

            // Output the subtitle
            Console.WriteLine("Chart subtitle: " + subtitleText);
        }
        else
        {
            Console.WriteLine("No charts found in the worksheet.");
        }
    }
}
