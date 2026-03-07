using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Timelines;

namespace AsposeCellsExamples
{
    public class ExportTimelineXlsx
    {
        public static void Run()
        {
            // Determine the path to the source workbook
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string sourcePath = Path.Combine(baseDir, "SourceWithTimeline.xlsx");

            // If the source file does not exist, create a placeholder workbook
            if (!File.Exists(sourcePath))
            {
                Workbook placeholder = new Workbook();
                placeholder.Worksheets[0].Name = "Sheet1";
                placeholder.Save(sourcePath, SaveFormat.Xlsx);
                Console.WriteLine($"Placeholder workbook created at \"{sourcePath}\".");
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Verify that the workbook contains at least one worksheet with a Timeline
            bool hasTimeline = false;
            foreach (Worksheet sheet in sourceWorkbook.Worksheets)
            {
                if (sheet.Timelines.Count > 0)
                {
                    hasTimeline = true;
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" contains {sheet.Timelines.Count} timeline(s).");
                }
            }

            if (!hasTimeline)
            {
                Console.WriteLine("No Timeline found in the source workbook.");
            }

            // Save the workbook (including its Timeline) to a new XLSX file
            string destPath = Path.Combine(baseDir, "ExportedTimeline.xlsx");
            sourceWorkbook.Save(destPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook with Timeline exported successfully to \"{destPath}\".");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportTimelineXlsx.Run();
        }
    }
}