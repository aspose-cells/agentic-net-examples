// Title: C# – Compare Aspose.Cells HTML Export With and Without Scientific Notation Suppression
// Description: Shows how to export a workbook to HTML using Aspose.Cells, first with default settings that may render numbers in scientific notation, then applying the custom format "0.####################" to force plain decimal output, and finally comparing the two HTML results.
// Keywords: Aspose.Cells | C# | HTML export | scientific notation | exponential notation | custom number format | HtmlSaveOptions | suppress exponential notation | compare HTML output | extract cell values
// Common Searches: Aspose.Cells prevent scientific notation in HTML | C# export Excel to HTML without exponential format | apply custom number format before HTML export Aspose.Cells | compare default and formatted HTML output Aspose.Cells | extract cell values from Aspose.Cells generated HTML
// Developer Intent: Verify that a custom number format eliminates scientific notation from the HTML produced by Aspose.Cells.
// Use Cases: Automated testing of numeric formatting in generated HTML reports | Creating financial or invoice HTML where large or tiny numbers must appear as plain decimals | Ensuring compliance‑driven presentations of exported data | Documenting exact numeric values without scientific notation | Adding a CI step to validate number‑formatting rules after HTML export
// AI Prompts: Write a C# method that loads two Aspose.Cells HTML files, extracts numeric cell strings, and returns true if they match after formatting. | Show how to apply a custom number format to an entire column or range in a workbook before exporting to HTML with Aspose.Cells. | Explain which HtmlSaveOptions properties preserve custom number formats during HTML export. | Create a script that logs differences between default and formatted HTML exports for all numeric cells. | Provide PowerShell commands to run the comparison program and capture its output for further analysis.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlComparison
{
    // Shows how to export a workbook to HTML using Aspose.Cells, first with default settings that may render numbers in scientific notation, then applying the custom format "0.####################" to force plain decimal output, and finally comparing the two HTML results.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add numbers that normally appear in exponential notation when exported
            // Large number
            sheet.Cells["A1"].PutValue(1234567890123456.0);
            // Small number
            sheet.Cells["A2"].PutValue(0.000000123456789);

            // -----------------------------------------------------------------
            // Export HTML with default settings (exponential notation may appear)
            // -----------------------------------------------------------------
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions();
            // Keep default ExportFormula value (true) – not relevant for this demo
            workbook.Save("default.html", defaultOptions);

            // -----------------------------------------------------------------
            // Apply a custom number format to suppress exponential notation
            // The format "0.####################" forces the value to be displayed
            // as a plain decimal string without scientific notation.
            // -----------------------------------------------------------------
            Style suppressExpStyle = workbook.CreateStyle();
            suppressExpStyle.Custom = "0.####################";

            // Apply the style to the cells containing the numbers
            sheet.Cells["A1"].SetStyle(suppressExpStyle);
            sheet.Cells["A2"].SetStyle(suppressExpStyle);

            // Export HTML after suppressing exponential notation
            HtmlSaveOptions suppressedOptions = new HtmlSaveOptions();
            workbook.Save("suppressed.html", suppressedOptions);

            // -----------------------------------------------------------------
            // Load the generated HTML files and display the numeric values
            // for manual verification of the difference.
            // -----------------------------------------------------------------
            string htmlDefault = File.ReadAllText("default.html");
            string htmlSuppressed = File.ReadAllText("suppressed.html");

            Console.WriteLine("=== Default HTML (may contain exponential notation) ===");
            Console.WriteLine(htmlDefault);
            Console.WriteLine();
            Console.WriteLine("=== Suppressed HTML (plain decimal notation) ===");
            Console.WriteLine(htmlSuppressed);
            Console.WriteLine();

            // Simple extraction of the cell values from the HTML for comparison
            // (Assumes the first <td> elements correspond to A1 and A2)
            string[] defaultValues = ExtractCellValues(htmlDefault);
            string[] suppressedValues = ExtractCellValues(htmlSuppressed);

            Console.WriteLine("Cell A1 value - Default:    " + defaultValues[0]);
            Console.WriteLine("Cell A1 value - Suppressed: " + suppressedValues[0]);
            Console.WriteLine("Cell A2 value - Default:    " + defaultValues[1]);
            Console.WriteLine("Cell A2 value - Suppressed: " + suppressedValues[1]);
        }

        // Helper method to extract the first two <td> contents from an HTML string.
        // This is a lightweight parser sufficient for the demonstration purpose.
        private static string[] ExtractCellValues(string html)
        {
            string[] result = new string[2] { string.Empty, string.Empty };
            int index = 0;
            int searchPos = 0;

            while (index < 2)
            {
                int tdStart = html.IndexOf("<td", searchPos, StringComparison.OrdinalIgnoreCase);
                if (tdStart == -1) break;

                int tdClose = html.IndexOf('>', tdStart);
                if (tdClose == -1) break;

                int tdEnd = html.IndexOf("</td>", tdClose, StringComparison.OrdinalIgnoreCase);
                if (tdEnd == -1) break;

                string cellContent = html.Substring(tdClose + 1, tdEnd - tdClose - 1).Trim();
                result[index] = System.Net.WebUtility.HtmlDecode(cellContent);
                index++;
                searchPos = tdEnd + 5; // Move past </td>
            }

            return result;
        }
    }
}
