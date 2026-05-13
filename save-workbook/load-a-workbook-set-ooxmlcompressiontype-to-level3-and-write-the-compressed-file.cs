using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCompressionDemo
{
    public class CompressionToStream
    {
        public static void Run()
        {
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            // CompressionType setting removed to use default compression
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.Save(stream, saveOptions);
                File.WriteAllBytes("compressed_output.xlsx", stream.ToArray());
            }
            workbook.Dispose();
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CompressionToStream.Run();
        }
    }
}