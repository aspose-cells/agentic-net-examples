using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWebApiDemo
{
    public class WorkbookMergeAndStream
    {
        public static MemoryStream GetMergedWorkbookStream(string firstFilePath, string secondFilePath)
        {
            Workbook firstWorkbook = File.Exists(firstFilePath) ? new Workbook(firstFilePath) : new Workbook();
            Workbook secondWorkbook = File.Exists(secondFilePath) ? new Workbook(secondFilePath) : new Workbook();

            firstWorkbook.Combine(secondWorkbook);

            var stream = new MemoryStream();
            firstWorkbook.Save(stream, new XlsSaveOptions());
            stream.Position = 0;
            return stream;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string firstFile = Path.Combine(baseDir, "first.xlsx");
            string secondFile = Path.Combine(baseDir, "second.xlsx");

            using var mergedStream = WorkbookMergeAndStream.GetMergedWorkbookStream(firstFile, secondFile);
            using var fileStream = File.Create(Path.Combine(baseDir, "merged.xls"));
            mergedStream.CopyTo(fileStream);
        }
    }
}