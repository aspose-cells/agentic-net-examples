using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWarningDemo
{
    public class CollectingWarningCallback : IWarningCallback
    {
        public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

        public void Warning(WarningInfo warningInfo)
        {
            Warnings.Add(warningInfo);
        }
    }

    class Program
    {
        static void Main()
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "sample_with_issues.xlsx");

            if (!File.Exists(filePath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Sample");
                tempWb.Save(filePath);
            }

            var warningCallback = new CollectingWarningCallback();

            var loadOptions = new LoadOptions
            {
                WarningCallback = warningCallback
            };

            var workbook = new Workbook(filePath, loadOptions);

            Console.WriteLine($"Total warnings captured during load: {warningCallback.Warnings.Count}");
            foreach (var warning in warningCallback.Warnings)
            {
                Console.WriteLine($"Warning Type: {warning.Type}");
                Console.WriteLine($"Description : {warning.Description}");
                Console.WriteLine(new string('-', 40));
            }

            Console.WriteLine($"Workbook loaded. Worksheets count: {workbook.Worksheets.Count}");
        }
    }
}