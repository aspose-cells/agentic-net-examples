using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToExcelWithStreamProvider
{
    public class HtmlReadStreamProvider : IStreamProvider
    {
        public void InitStream(StreamProviderOptions options)
        {
            options.Stream = File.OpenRead(options.DefaultPath);
        }

        public void CloseStream(StreamProviderOptions options)
        {
            if (options.Stream != null)
            {
                options.Stream.Close();
                options.Stream = null;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            string htmlFilePath = "input.html";
            string excelFilePath = "output.xlsx";

            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                StreamProvider = new HtmlReadStreamProvider()
            };

            Workbook workbook = new Workbook(htmlFilePath, loadOptions);
            workbook.Save(excelFilePath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML file '{htmlFilePath}' has been converted to Excel file '{excelFilePath}'.");
        }
    }
}