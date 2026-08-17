// Title: Bulk Convert Excel Dates to Japanese Era (Heisei, Reiwa) with Aspose.Cells for .NET
// Description: Loads a workbook, sets its region to Japan, defines a custom "ggge" number format for Japanese eras, creates a style, scans the used range and applies the style to every DateTime cell, then saves the file so dates appear as "平成31年4月30日" or "令和1年5月1日".
// Keywords: Aspose.Cells | C# | Japanese era formatting | ggge pattern | bulk date conversion | Excel region Japan | locale-aware dates | Heisei | Reiwa | globalization | localization
// Common Searches: Aspose.Cells format dates with Japanese era | C# apply Japanese era style to all Excel date cells | set workbook region to Japan Aspose.Cells | ggge custom number format example | convert Gregorian dates to Japanese era in .NET
// Developer Intent: Automatically apply the appropriate Japanese era format to every DateTime cell during bulk Excel processing.
// Use Cases: Localize financial reports by converting Gregorian dates to era notation across an entire workbook. | Prepare a Japanese‑language version of a template where all date columns must display Heisei or Reiwa eras. | Run a batch job that standardizes date formatting for multiple Excel files before distribution to Japanese stakeholders.
// AI Prompts: Generate C# code using Aspose.Cells that sets the workbook region to Japan and applies a "ggge" era format to all DateTime cells. | Explain how the "ggge" pattern selects the correct Japanese era based on the year in Aspose.Cells custom formats. | Provide troubleshooting steps when Japanese era formatting does not appear after applying the custom style.

using System;
using Aspose.Cells;
using Aspose.Cells.Settings;

namespace AsposeCellsJapaneseEraDemo
{
    // Loads a workbook, sets its region to Japan, defines a custom "ggge" number format for Japanese eras, creates a style, scans the used range and applies the style to every DateTime cell, then saves the file so dates appear as "平成31年4月30日" or "令和1年5月1日".
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (bulk conversion source)
            // Replace "input.xlsx" with the path to your source workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Set the workbook region to Japan so that Japanese era formatting is recognized
            workbook.Settings.Region = CountryCode.Japan;

            // Define a custom number format that displays dates using Japanese eras
            // Example format: "平成31年4月30日" or "令和1年5月1日"
            // The format string uses the "ggge" pattern for era name and year.
            string japaneseEraFormat = "[$-ja-JP]ggge\"年\"M\"月\"d\"日\"";

            // Create a style with the custom format
            Style eraStyle = workbook.CreateStyle();
            eraStyle.Custom = japaneseEraFormat;

            // Apply the style to all cells that contain DateTime values
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Iterate through used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Type == CellValueType.IsDateTime)
                    {
                        // Apply the Japanese era style
                        cell.SetStyle(eraStyle);
                    }
                }
            }

            // Save the workbook with converted dates
            // Replace "output.xlsx" with the desired output path
            workbook.Save("output.xlsx");
        }
    }
}
