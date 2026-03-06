using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    // Demonstrates various XLSX related scenarios using Aspose.Cells for .NET
    public static class XlsxUsageScenarios
    {
        // 1. Create a new workbook and save it as XLSX
        public static void CreateAndSaveXlsx()
        {
            Workbook workbook = new Workbook();

            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Employees";
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(28);

            workbook.Save("Employees.xlsx", SaveFormat.Xlsx);
        }

        // 2. Load an existing XLSX file from disk
        public static void LoadXlsxFromFile()
        {
            Workbook workbook = new Workbook("Employees.xlsx");

            Worksheet sheet = workbook.Worksheets["Employees"];
            string name = sheet.Cells["A2"].StringValue;
            int age = (int)sheet.Cells["B2"].IntValue;

            Console.WriteLine($"Loaded row: {name}, {age}");
        }

        // 3. Load an XLSX file from a MemoryStream
        public static void LoadXlsxFromStream()
        {
            byte[] fileBytes = File.ReadAllBytes("Employees.xlsx");

            using (MemoryStream ms = new MemoryStream(fileBytes))
            {
                Workbook workbook = new Workbook(ms);

                workbook.Worksheets[0].Cells["C1"].PutValue("Department");
                workbook.Worksheets[0].Cells["C2"].PutValue("Sales");

                workbook.Save("Employees_Modified.xlsx", SaveFormat.Xlsx);
            }
        }

        // 4. Convert an XLSX file to another format (e.g., PDF) using ConversionUtility
        public static void ConvertXlsxToPdf()
        {
            string source = "Employees.xlsx";
            string destination = "Employees.pdf";

            ConversionUtility.Convert(source, destination);
        }

        // 5. Detect the format of a file (should report XLSX) using FileFormatUtil
        public static void DetectFileFormat()
        {
            string path = "Employees.xlsx";

            FileFormatInfo info = FileFormatUtil.DetectFileFormat(path);
            Console.WriteLine($"Detected format: {info.FileFormatType}");
            Console.WriteLine($"Load format: {info.LoadFormat}");
        }

        // 6. Save an XLSX workbook with custom XlsSaveOptions (demonstrates using options even though format is XLSX)
        public static void SaveWithXlsOptions()
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");

            XlsSaveOptions options = new XlsSaveOptions
            {
                MatchColor = true,
                ClearData = false
            };

            workbook.Save("Sample.xls", options);
        }

        // 7. Export data to CSV while keeping the source workbook as XLSX
        public static void ExportXlsxToCsv()
        {
            Workbook workbook = new Workbook("Employees.xlsx");
            workbook.Save("Employees.csv", SaveFormat.Csv);
        }

        // 8. Use LoadOptions to explicitly specify XLSX format when loading
        public static void LoadWithExplicitLoadOptions()
        {
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook("Employees.xlsx", loadOptions);
            Console.WriteLine($"Worksheets count: {workbook.Worksheets.Count}");
        }

        // Entry point to run all scenarios
        public static void RunAll()
        {
            CreateAndSaveXlsx();
            LoadXlsxFromFile();
            LoadXlsxFromStream();
            ConvertXlsxToPdf();
            DetectFileFormat();
            SaveWithXlsOptions();
            ExportXlsxToCsv();
            LoadWithExplicitLoadOptions();
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            XlsxUsageScenarios.RunAll();
        }
    }
}