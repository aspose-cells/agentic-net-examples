// Title: Copy a Range to a New Workbook and Set Worksheet Tab Color by Content Type (Aspose.Cells C#)
// Description: Creates a source workbook, fills A1:C3 with text, numbers, dates and formulas, copies the range to a destination workbook, analyzes the source cells to detect data types, and applies a tab color (purple, green, yellow, orange) to the destination sheet based on the highest‑priority type before saving the file.
// Keywords: Aspose.Cells copy range | C# copy range to new workbook | set worksheet tab color programmatically | detect formulas Aspose.Cells | range content type detection | tab color based on data type | Aspose.Cells Range.Copy example | Excel automation Aspose.Cells
// Common Searches: Aspose.Cells copy range and change tab color | C# set worksheet tab color after copying range | detect numeric or formula cells in Aspose.Cells | how to color Excel sheet tab by content type using Aspose | copy range to another workbook Aspose.Cells C#
// Developer Intent: Copy a defined cell range to a new workbook and automatically color the destination worksheet tab according to the data types present in the source range.
// Use Cases: Generate summary reports where the tab color instantly signals the presence of formulas, helping reviewers locate calculated sheets. | Create financial workbooks that highlight numeric‑only sheets with a green tab for quick identification of data tables. | Produce multi‑sheet exports where each sheet’s tab color reflects its dominant content (text, date, formula) to improve navigation for end users.
// AI Prompts: Write C# code with Aspose.Cells that copies a range and sets the destination worksheet tab color based on whether the range contains formulas, numbers, strings, or dates. | Explain how to modify the priority order of tab colors when a range includes multiple data types in Aspose.Cells. | Suggest a pattern to apply different tab colors to several destination worksheets, each reflecting the content type of its own copied range.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    // Creates a source workbook, fills A1:C3 with text, numbers, dates and formulas, copies the range to a destination workbook, analyzes the source cells to detect data types, and applies a tab color (purple, green, yellow, orange) to the destination sheet based on the highest‑priority type before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create source workbook and populate a range ----------
                Workbook sourceWb = new Workbook();                     // create source workbook
                Worksheet sourceWs = sourceWb.Worksheets[0];            // get first worksheet
                sourceWs.Name = "Source";

                // Fill a sample range A1:C3 with mixed data types
                sourceWs.Cells["A1"].PutValue("Text");                 // string
                sourceWs.Cells["B1"].PutValue(123);                    // numeric
                sourceWs.Cells["C1"].PutValue(DateTime.Now);           // date/time
                sourceWs.Cells["A2"].PutValue("More Text");
                sourceWs.Cells["B2"].PutValue(456);
                sourceWs.Cells["C2"].PutValue("=SUM(B1,B2)");          // formula
                sourceWs.Cells["A3"].PutValue(789);
                sourceWs.Cells["B3"].PutValue("End");
                sourceWs.Cells["C3"].PutValue(3.14);

                // Define the source range to copy
                AsposeRange sourceRange = sourceWs.Cells.CreateRange("A1:C3");

                // ---------- Create destination workbook ----------
                Workbook destWb = new Workbook();                       // create destination workbook
                Worksheet destWs = destWb.Worksheets[0];                // get first worksheet
                destWs.Name = "Destination";

                // Define the destination range (same size, starting at A1)
                AsposeRange destRange = destWs.Cells.CreateRange("A1:C3");

                // ---------- Copy the range ----------
                // The Copy method copies data, formulas, formatting, etc.
                sourceRange.Copy(destRange);

                // ---------- Determine content type of the source range ----------
                bool hasNumeric = false;
                bool hasString = false;
                bool hasDate = false;
                bool hasFormula = false;

                // Iterate through each cell in the source range
                for (int row = sourceRange.FirstRow; row <= sourceRange.FirstRow + sourceRange.RowCount - 1; row++)
                {
                    for (int col = sourceRange.FirstColumn; col <= sourceRange.FirstColumn + sourceRange.ColumnCount - 1; col++)
                    {
                        Cell cell = sourceWs.Cells[row, col];
                        if (cell.IsFormula)
                        {
                            hasFormula = true;
                        }
                        else if (cell.Type == CellValueType.IsNumeric)
                        {
                            hasNumeric = true;
                        }
                        else if (cell.Type == CellValueType.IsString)
                        {
                            hasString = true;
                        }
                        else if (cell.Type == CellValueType.IsDateTime)
                        {
                            hasDate = true;
                        }
                    }
                }

                // ---------- Set worksheet tab color based on detected content ----------
                // Priority: Formula > Numeric > String > Date > Default
                if (hasFormula)
                    destWs.TabColor = Color.Purple;
                else if (hasNumeric)
                    destWs.TabColor = Color.Green;
                else if (hasString)
                    destWs.TabColor = Color.Yellow;
                else if (hasDate)
                    destWs.TabColor = Color.Orange;
                else
                    destWs.TabColor = Color.Empty; // no specific color

                // ---------- Save the destination workbook ----------
                destWb.Save("RangeCopyWithTabColor.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
