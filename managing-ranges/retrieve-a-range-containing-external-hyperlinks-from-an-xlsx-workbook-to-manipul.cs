using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalHyperlinkDemo
{
    public class Program
    {
        public static void Main()
        {
            string inputPath = "InputWithHyperlinks.xlsx";
            Workbook workbook;

            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Sheet1";

                // Add sample data and a hyperlink
                ws.Cells["A1"].PutValue("Google");
                ws.Hyperlinks.Add(0, 0, 1, 1, "https://www.google.com");

                workbook.Save(inputPath);
            }

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                HyperlinkCollection hyperlinks = sheet.Hyperlinks;

                foreach (Hyperlink link in hyperlinks)
                {
                    CellArea area = link.Area;

                    int rowCount = area.EndRow - area.StartRow + 1;
                    int columnCount = area.EndColumn - area.StartColumn + 1;

                    Aspose.Cells.Range hyperlinkRange = sheet.Cells.CreateRange(
                        area.StartRow,
                        area.StartColumn,
                        rowCount,
                        columnCount);

                    Style style = workbook.CreateStyle();
                    style.ForegroundColor = Color.LightYellow;
                    style.Pattern = BackgroundType.Solid;

                    StyleFlag flag = new StyleFlag { CellShading = true };
                    hyperlinkRange.ApplyStyle(style, flag);

                    hyperlinkRange[0, 0].PutValue("Linked Cell");
                }
            }

            workbook.Save("OutputWithModifiedHyperlinkRanges.xlsx");
        }
    }
}