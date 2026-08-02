// Title: Flag Formatting‑Only Worksheets in Aspose.Cells for .NET (C#) Using MaxDataRow & Row.IsBlank
// Description: C# sample that creates or loads a workbook, iterates through each worksheet, uses Cells.MaxDataRow and Row.IsBlank to detect sheets that contain only formatting (no cell values), renames them with a "FLAGGED_" prefix, and saves the result.
// Keywords: Aspose.Cells C# flag formatting only worksheet | MaxDataRow blank rows detection | Row.IsBlank Aspose.Cells | rename worksheet without data | detect empty data sheets .NET | formatting‑only sheet detection | Aspose.Cells worksheet validation | C# workbook cleanup
// Common Searches: Aspose.Cells detect formatting only sheet | C# check if worksheet has only formatting | MaxDataRow returns rows with formatting only | rename empty worksheets Aspose.Cells | how to flag blank rows in Aspose.Cells
// Developer Intent: Identify worksheets whose MaxDataRow is greater than zero yet contain no actual cell values, and rename them to indicate they are formatting‑only.
// Use Cases: Automatically mark layout‑only sheets before distributing a workbook to end users. | Skip processing of worksheets that appear to have data rows but are actually empty, improving batch report generation. | Maintain a clean workbook by flagging sheets created solely for visual design. | Generate audit logs of formatting‑only worksheets for compliance reporting.
// AI Prompts: Generate C# code with Aspose.Cells that flags worksheets where MaxDataRow > 0 but every row is blank, adding a "FLAGGED_" prefix to the sheet name. | Suggest a performance‑optimized method to detect formatting‑only worksheets without iterating every row in Aspose.Cells for .NET. | Create a reusable Aspose.Cells utility method that returns a list of worksheet names containing only formatting based on MaxDataRow and Row.IsBlank.

using System;
using Aspose.Cells;

namespace WorksheetFlaggingDemo
{
    // C# sample that creates or loads a workbook, iterates through each worksheet, uses Cells.MaxDataRow and Row.IsBlank to detect sheets that contain only formatting (no cell values), renames them with a "FLAGGED_" prefix, and saves the result.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook or load an existing one
            // Replace the path with your actual file if loading
            Workbook workbook = new Workbook(); // new workbook for demonstration

            // Example: add a worksheet with formatting only (no data)
            Worksheet fmtOnlySheet = workbook.Worksheets[0];
            fmtOnlySheet.Name = "FormattingOnlySheet";
            // Apply some formatting without putting any values
            Style style = workbook.CreateStyle();
            style.Font.IsBold = true;
            fmtOnlySheet.Cells.CreateRange("A1:C5").ApplyStyle(style, new StyleFlag { FontBold = true });

            // Example: add a worksheet with actual data
            Worksheet dataSheet = workbook.Worksheets.Add("DataSheet");
            dataSheet.Cells["A1"].PutValue("Header");
            dataSheet.Cells["A2"].PutValue("Value");

            // Iterate through all worksheets to flag those that have MaxDataRow > 0
            // but contain no actual data (all rows are blank, i.e., formatting‑only)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int maxDataRow = sheet.Cells.MaxDataRow;

                // MaxDataRow > 0 indicates there is at least one row index considered
                if (maxDataRow > 0)
                {
                    bool allRowsBlank = true;

                    // Check each row up to MaxDataRow using Row.IsBlank
                    for (int rowIndex = 0; rowIndex <= maxDataRow; rowIndex++)
                    {
                        Row row = sheet.Cells.Rows[rowIndex];
                        if (!row.IsBlank)
                        {
                            allRowsBlank = false;
                            break;
                        }
                    }

                    // If all rows are blank, rename the worksheet to flag it
                    if (allRowsBlank)
                    {
                        sheet.Name = "FLAGGED_" + sheet.Name;
                    }
                }
            }

            // Save the workbook
            workbook.Save("FlaggedWorksheets.xlsx", SaveFormat.Xlsx);
        }
    }
}
