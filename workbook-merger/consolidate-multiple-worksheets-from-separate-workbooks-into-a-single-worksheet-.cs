using System;
using Aspose.Cells;

public class WorksheetConsolidator
{
    public static void ConsolidateWorksheets(string[] sourceFiles, string targetFilePath)
    {
        using (Workbook targetWorkbook = new Workbook())
        {
            Worksheet consolidatedSheet = targetWorkbook.Worksheets.Add("Consolidated");
            int nextRow = 0;

            foreach (string srcPath in sourceFiles)
            {
                using (Workbook srcWorkbook = new Workbook(srcPath))
                {
                    foreach (Worksheet srcSheet in srcWorkbook.Worksheets)
                    {
                        int maxRow = srcSheet.Cells.MaxDataRow;
                        int maxCol = srcSheet.Cells.MaxDataColumn;

                        for (int r = 0; r <= maxRow; r++)
                        {
                            for (int c = 0; c <= maxCol; c++)
                            {
                                Cell srcCell = srcSheet.Cells[r, c];
                                if (srcCell.Value != null)
                                {
                                    int destRow = nextRow + r;
                                    int destCol = c;

                                    Cell destCell = consolidatedSheet.Cells[destRow, destCol];
                                    destCell.PutValue(srcCell.Value);
                                    destCell.SetStyle(srcCell.GetStyle());
                                }
                            }
                        }

                        nextRow += maxRow + 2; // blank row between worksheets
                    }
                }
            }

            targetWorkbook.Save(targetFilePath, SaveFormat.Xlsx);
        }
    }

    public static void Main()
    {
        string[] sourceFiles = new string[]
        {
            "Source1.xlsx",
            "Source2.xlsx",
            "Source3.xlsx"
        };

        string outputPath = "ConsolidatedWorkbook.xlsx";

        ConsolidateWorksheets(sourceFiles, outputPath);

        Console.WriteLine($"Worksheets have been consolidated into '{outputPath}'.");
    }
}