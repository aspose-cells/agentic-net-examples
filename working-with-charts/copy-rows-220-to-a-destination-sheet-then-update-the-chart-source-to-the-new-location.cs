// Title: Copy rows 2‑20 from a source worksheet to another sheet and update chart series source with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy a range of rows to a different worksheet while preserving formulas | How to move rows 2 to 20 to another sheet and change the chart data source using Aspose.Cells | Preserve cell styles when transferring rows between worksheets in Aspose.Cells for .NET | Update chart series range to point to a new worksheet after copying data with Aspose.Cells | Extract column letter from a chart formula in Aspose.Cells C# example
// Tags: transfer row range between worksheets Aspose.Cells C# | reassign chart series data range Aspose.Cells | retain cell formatting and formulas during copy Aspose.Cells | parse column identifier from chart formula Aspose.Cells | manual cell‑by‑cell copy using Cells collection Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample loads or creates a workbook, ensures "Source" and "Destination" worksheets exist, copies rows 2‑20 (zero‑based 1‑19) from the source sheet to the top of the destination sheet while preserving values, styles, and formulas, then locates the first chart on the source sheet and rewrites its series Values and XValues formulas to reference the newly copied range on the Destination sheet using a helper that extracts column letters from the original formulas, and finally saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load workbook if it exists; otherwise create a new one with required sheets.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets.Add("Source");
                workbook.Worksheets.Add("Destination");
            }

            // Ensure source and destination worksheets exist.
            Worksheet srcSheet = workbook.Worksheets["Source"] ?? workbook.Worksheets.Add("Source");
            Worksheet destSheet = workbook.Worksheets["Destination"] ?? workbook.Worksheets.Add("Destination");

            // Define rows to copy (rows 2‑20 => zero‑based indices 1‑19)
            int srcStartRow = 1;          // row 2
            int srcEndRow = 19;           // row 20
            int rowCount = srcEndRow - srcStartRow + 1;
            int destStartRow = 0;         // copy to row 1 in destination sheet

            // Copy rows from source sheet to destination sheet manually
            try
            {
                int maxColumn = srcSheet.Cells.MaxColumn + 1;
                for (int i = 0; i < rowCount; i++)
                {
                    int srcRow = srcStartRow + i;
                    int destRow = destStartRow + i;
                    for (int col = 0; col < maxColumn; col++)
                    {
                        Cell srcCell = srcSheet.Cells[srcRow, col];
                        Cell destCell = destSheet.Cells[destRow, col];

                        // Copy value
                        destCell.PutValue(srcCell.Value);

                        // Copy style
                        destCell.SetStyle(srcCell.GetStyle());

                        // Copy formula if present
                        if (!string.IsNullOrEmpty(srcCell.Formula))
                        {
                            destCell.Formula = srcCell.Formula;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Row copy error: {ex.Message}");
            }

            // Update chart source to point to the new location (assumes chart is on source sheet)
            try
            {
                if (srcSheet.Charts.Count > 0)
                {
                    Chart chart = srcSheet.Charts[0];

                    for (int s = 0; s < chart.NSeries.Count; s++)
                    {
                        dynamic series = chart.NSeries[s];

                        // Update Values formula
                        if (series.Values != null && series.Values.Length > 0)
                        {
                            string originalFormula = series.Values[0];
                            string column = ExtractColumnFromFormula(originalFormula);
                            string newFormula = $"=Destination!${column}$1:${column}${rowCount}";
                            series.Values = new[] { newFormula };
                        }

                        // Update XValues formula if present
                        if (series.XValues != null && series.XValues.Length > 0)
                        {
                            string xOriginal = series.XValues[0];
                            string xColumn = ExtractColumnFromFormula(xOriginal);
                            string xNew = $"=Destination!${xColumn}$1:${xColumn}${rowCount}";
                            series.XValues = new[] { xNew };
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Chart update error: {ex.Message}");
            }

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Extract column letters (e.g., "B") from a formula like =Sheet!$B$2:$B$20
    static string ExtractColumnFromFormula(string formula)
    {
        if (string.IsNullOrEmpty(formula))
            return "A";

        int exclPos = formula.IndexOf('!');
        if (exclPos < 0) exclPos = -1;

        int firstDollar = formula.IndexOf('$', exclPos);
        if (firstDollar < 0) return "A";

        int secondDollar = formula.IndexOf('$', firstDollar + 1);
        if (secondDollar < 0) secondDollar = formula.Length;

        return formula.Substring(firstDollar + 1, secondDollar - firstDollar - 1);
    }
}
