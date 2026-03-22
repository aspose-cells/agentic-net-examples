using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTimelineExport
{
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportTimeline.Run();
        }
    }

    public class ExportTimeline
    {
        public static void Run()
        {
            const string sourcePath = "SourceTemplate.xltm";
            const string outputPath = "ExportedTimeline.xlsx";

            // Ensure the source template exists; create a minimal one if it does not.
            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                tempWb.Save(sourcePath, SaveFormat.Xltm);
            }

            // Load the source XLTM workbook (macro‑enabled template) which may contain Timeline controls
            var sourceWorkbook = new Workbook(sourcePath);

            // Save the workbook as a regular XLSX file.
            // All worksheet content, including Timelines, is preserved during the conversion.
            sourceWorkbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}