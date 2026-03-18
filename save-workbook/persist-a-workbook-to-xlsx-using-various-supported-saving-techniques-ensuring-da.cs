using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSavingTechniques
{
    public class WorkbookSavingDemo
    {
        public static void Run()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and add sample data
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // ------------------------------------------------------------
            // 2. Simple save using file name and explicit SaveFormat (XLSX)
            // ------------------------------------------------------------
            string pathExplicitXlsx = "ExplicitSave.xlsx";
            workbook.Save(pathExplicitXlsx, SaveFormat.Xlsx);

            // Verify data integrity by reloading
            Workbook verifyExplicit = new Workbook(pathExplicitXlsx);
            Console.WriteLine("Explicit Save - Cell B2: " + verifyExplicit.Worksheets[0].Cells["B2"].StringValue);

            // ------------------------------------------------------------
            // 3. Save using only file name (format inferred from extension)
            // ------------------------------------------------------------
            string pathAutoXlsx = "AutoDetectedSave.xlsx";
            workbook.Save(pathAutoXlsx); // extension .xlsx triggers Xlsx format

            // Verify data integrity
            Workbook verifyAuto = new Workbook(pathAutoXlsx);
            Console.WriteLine("Auto Save - Cell B3: " + verifyAuto.Worksheets[0].Cells["B3"].StringValue);

            // ------------------------------------------------------------
            // 4. Save with OoxmlSaveOptions (e.g., enable compression, export cell names)
            // ------------------------------------------------------------
            string pathOoxmlOptions = "OoxmlOptionsSave.xlsx";
            OoxmlSaveOptions ooxmlOptions = new OoxmlSaveOptions
            {
                CompressionType = OoxmlCompressionType.Level6,
                ExportCellName = true,
                UpdateZoom = true
            };
            workbook.Save(pathOoxmlOptions, ooxmlOptions);

            // Verify data integrity
            Workbook verifyOoxml = new Workbook(pathOoxmlOptions);
            Console.WriteLine("Ooxml Options Save - Cell A1: " + verifyOoxml.Worksheets[0].Cells["A1"].StringValue);

            // ------------------------------------------------------------
            // 5. Save to a MemoryStream using Save(Stream, SaveFormat) and write to file
            // ------------------------------------------------------------
            string pathStreamXlsx = "StreamSave.xlsx";
            using (MemoryStream ms = new MemoryStream())
            {
                // Save workbook into the stream in XLSX format
                workbook.Save(ms, SaveFormat.Xlsx);

                // Reset stream position before reading
                ms.Position = 0;

                // Write stream content to a physical file
                using (FileStream file = new FileStream(pathStreamXlsx, FileMode.Create, FileAccess.Write))
                {
                    ms.CopyTo(file);
                }
            }

            // Verify data integrity
            Workbook verifyStream = new Workbook(pathStreamXlsx);
            Console.WriteLine("Stream Save - Cell A2: " + verifyStream.Worksheets[0].Cells["A2"].IntValue);

            // ------------------------------------------------------------
            // 6. Save to a MemoryStream using Save(Stream, SaveOptions) with XlsSaveOptions (XLS format)
            // ------------------------------------------------------------
            string pathStreamXls = "StreamSave.xls";
            XlsSaveOptions xlsOptions = new XlsSaveOptions
            {
                MatchColor = true,
                ClearData = false,
                ValidateMergedAreas = true,
                MergeAreas = true,
                SortNames = true,
                RefreshChartCache = true
            };
            using (MemoryStream msXls = new MemoryStream())
            {
                // Save workbook into the stream in XLS format using options
                workbook.Save(msXls, xlsOptions);

                // Reset stream position before writing to file
                msXls.Position = 0;

                // Write stream content to a physical file
                using (FileStream fileXls = new FileStream(pathStreamXls, FileMode.Create, FileAccess.Write))
                {
                    msXls.CopyTo(fileXls);
                }
            }

            // Verify data integrity
            Workbook verifyStreamXls = new Workbook(pathStreamXls);
            Console.WriteLine("Stream XLS Save - Cell B2: " + verifyStreamXls.Worksheets[0].Cells["B2"].StringValue);

            // ------------------------------------------------------------
            // Cleanup
            // ------------------------------------------------------------
            workbook.Dispose();
            verifyExplicit.Dispose();
            verifyAuto.Dispose();
            verifyOoxml.Dispose();
            verifyStream.Dispose();
            verifyStreamXls.Dispose();

            Console.WriteLine("All saving techniques completed successfully.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookSavingDemo.Run();
        }
    }
}