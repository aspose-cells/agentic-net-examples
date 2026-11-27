// Title: Detect cells containing a leading apostrophe and rich‑text formatting using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells find cells with QuotePrefix and character-level formatting | list worksheet cells that have both a leading apostrophe and rich text in .NET | how to iterate the used range and check for apostrophe prefix and rich text using Aspose.Cells | identify Excel cells with quote prefix and bold or italic styling via Aspose.Cells API | scan an Excel workbook for cells containing an apostrophe and styled text in C#
// Tags: quote prefix detection Aspose.Cells C# | rich‑text character formatting check Aspose.Cells | used range iteration Aspose.Cells worksheets | cell address logging for apostrophe and styled text | detect combined apostrophe and rich text in Excel cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace DetectApostropheRichTextCells
{
    // The example loads a workbook (or creates a demo workbook), iterates over each worksheet's used range, and for each cell checks the Style.QuotePrefix flag to see if a leading apostrophe is present. It then examines the cell's Characters font properties (bold, italic, underline, color) to determine if rich‑text formatting exists. When both conditions are met, the cell's address (e.g., Sheet1!A1) is written to the console. If no such cells are found, a message is displayed.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Workbook workbook = null;
                string sourcePath = null;

                // Determine source workbook
                if (args.Length > 0 && File.Exists(args[0]))
                {
                    sourcePath = args[0];
                    Console.WriteLine($"Loading workbook from \"{sourcePath}\"...");
                    try
                    {
                        workbook = new Workbook(sourcePath);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Failed to load workbook: {ex.Message}");
                        return;
                    }
                }
                else
                {
                    Console.WriteLine("No valid input file supplied. Creating a demo workbook in memory.");
                    try
                    {
                        workbook = CreateDemoWorkbook();
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Failed to create demo workbook: {ex.Message}");
                        return;
                    }
                }

                if (workbook == null)
                {
                    Console.Error.WriteLine("Workbook is not initialized.");
                    return;
                }

                Console.WriteLine("Scanning worksheets for cells that have both a leading apostrophe and rich‑text formatting...");

                bool anyFound = false;

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Determine the used range to avoid iterating over the entire grid
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            if (cell == null)
                                continue;

                            // QuotePrefix indicates a leading apostrophe (')
                            bool hasLeadingApostrophe = cell.GetStyle().QuotePrefix;

                            // Detect rich‑text by checking if any character‑level formatting differs from default
                            bool hasRichText = false;
                            string text = cell.StringValue;
                            if (!string.IsNullOrEmpty(text))
                            {
                                Font font = cell.Characters(0, text.Length).Font;
                                if (font.IsBold ||
                                    font.IsItalic ||
                                    font.Underline != FontUnderlineType.None ||
                                    font.Color != Color.Empty)
                                {
                                    hasRichText = true;
                                }
                            }

                            if (hasLeadingApostrophe && hasRichText)
                            {
                                anyFound = true;
                                Console.WriteLine($"{sheet.Name}!{cell.Name}");
                            }
                        }
                    }
                }

                if (!anyFound)
                {
                    Console.WriteLine("No cells with both a leading apostrophe and rich‑text were found.");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }

        private static Workbook CreateDemoWorkbook()
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Cell A1 – leading apostrophe and rich text (bold)
            Cell a1 = ws.Cells["A1"];
            a1.PutValue("'Demo"); // The leading apostrophe is part of the string
            Style a1Style = a1.GetStyle();
            a1Style.QuotePrefix = true;
            a1.SetStyle(a1Style);
            a1.Characters(0, a1.StringValue.Length).Font.IsBold = true;

            // Cell B1 – leading apostrophe only
            Cell b1 = ws.Cells["B1"];
            b1.PutValue("'OnlyApostrophe");
            Style b1Style = b1.GetStyle();
            b1Style.QuotePrefix = true;
            b1.SetStyle(b1Style);

            // Cell C1 – rich‑text only (italic)
            Cell c1 = ws.Cells["C1"];
            c1.PutValue("RichOnly");
            c1.Characters(0, c1.StringValue.Length).Font.IsItalic = true;

            // Cell D1 – plain text
            ws.Cells["D1"].PutValue("Plain");

            // Save the demo workbook next to the executable for visual verification (optional)
            string demoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DemoWorkbook.xlsx");
            try
            {
                wb.Save(demoPath);
                Console.WriteLine($"Demo workbook saved to \"{demoPath}\".");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to save demo workbook: {ex.Message}");
            }

            return wb;
        }
    }
}
