using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Paths of the workbooks to be merged (first workbook will be the destination)
            string[] workbookFiles = { "Workbook1.xlsx", "Workbook2.xlsx", "Workbook3.xlsx" };

            // Load the first workbook – it will receive the merged content
            Workbook mergedWorkbook = new Workbook(workbookFiles[0]);

            // Iterate over the remaining workbooks and merge them one by one
            for (int i = 1; i < workbookFiles.Length; i++)
            {
                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(workbookFiles[i]);

                // Remember the current number of worksheets before combining
                int originalSheetCount = mergedWorkbook.Worksheets.Count;

                // Combine the source workbook into the destination workbook
                mergedWorkbook.Combine(sourceWorkbook);

                // After Combine, the new worksheets are appended at the end.
                // Their starting index is the original sheet count.
                int newSheetStartIndex = originalSheetCount;

                // Copy PageSetup settings from each source worksheet to the corresponding
                // newly added worksheet in the merged workbook.
                for (int j = 0; j < sourceWorkbook.Worksheets.Count; j++)
                {
                    Worksheet destSheet = mergedWorkbook.Worksheets[newSheetStartIndex + j];
                    Worksheet srcSheet = sourceWorkbook.Worksheets[j];

                    // Use PageSetup.Copy with default CopyOptions to preserve layout
                    destSheet.PageSetup.Copy(srcSheet.PageSetup, new CopyOptions());
                }
            }

            // Save the merged workbook with all original page‑setup settings preserved
            mergedWorkbook.Save("MergedWithPageSetup.xlsx", SaveFormat.Xlsx);
        }
    }
}