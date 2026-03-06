using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadingScenarios
{
    class Program
    {
        static void Main()
        {
            // Prepare sample files used in the scenarios
            CreateSampleFile("Sample1.xlsx");
            CreateSampleFile("Sample2.xlsx");
            CreateSampleFile("Sample3.xlsx");
            CreateMacroEnabledFile("MacroEnabled.xlsm");

            // Scenario 1: Load workbook from a file path using the default constructor.
            Workbook wbFromFile = new Workbook("Sample1.xlsx");
            Console.WriteLine("Scenario 1 - Worksheets count: " + wbFromFile.Worksheets.Count);
            wbFromFile.Save("Output_Scenario1.xlsx", SaveFormat.Xlsx);

            // Scenario 2: Load workbook from a file path with custom LoadOptions.
            LoadOptions options2 = new LoadOptions
            {
                ParsingFormulaOnOpen = false,
                IgnoreUselessShapes = true
            };
            Workbook wbWithOptions = new Workbook("Sample2.xlsx", options2);
            Console.WriteLine("Scenario 2 - ParsingFormulaOnOpen: " + options2.ParsingFormulaOnOpen);
            Console.WriteLine("Scenario 2 - IgnoreUselessShapes: " + options2.IgnoreUselessShapes);
            wbWithOptions.Save("Output_Scenario2.xlsx", SaveFormat.Xlsx);

            // Scenario 3: Load workbook from a memory stream using the Stream constructor.
            using (MemoryStream memStream = new MemoryStream())
            {
                Workbook tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Data from stream");
                tempWb.Save(memStream, SaveFormat.Xlsx);
                memStream.Position = 0;

                Workbook wbFromStream = new Workbook(memStream);
                Console.WriteLine("Scenario 3 - Cell A1 value: " + wbFromStream.Worksheets[0].Cells["A1"].StringValue);
                wbFromStream.Save("Output_Scenario3.xlsx", SaveFormat.Xlsx);
            }

            // Scenario 4: Load workbook from a stream with LoadOptions.
            using (MemoryStream memStream2 = new MemoryStream())
            {
                Workbook tempWb2 = new Workbook();
                tempWb2.Worksheets[0].Cells["B2"].PutValue(12345);
                tempWb2.Save(memStream2, SaveFormat.Xlsx);
                memStream2.Position = 0;

                LoadOptions options4 = new LoadOptions(LoadFormat.Xlsx)
                {
                    ParsingFormulaOnOpen = true
                };
                Workbook wbFromStreamWithOptions = new Workbook(memStream2, options4);
                Console.WriteLine("Scenario 4 - Cell B2 value: " + wbFromStreamWithOptions.Worksheets[0].Cells["B2"].IntValue);
                wbFromStreamWithOptions.Save("Output_Scenario4.xlsx", SaveFormat.Xlsx);
            }

            // Scenario 5: Copy a loaded workbook into a new workbook using Copy method.
            Workbook sourceWb = new Workbook("Sample3.xlsx");
            Workbook destWb = new Workbook();
            sourceWb.Copy(destWb);
            destWb.Save("Output_Scenario5_Copy.xlsx", SaveFormat.Xlsx);

            // Scenario 6: Copy with CopyOptions to keep macros (if any).
            Workbook sourceWithMacro = new Workbook("MacroEnabled.xlsm");
            Workbook destWithMacro = new Workbook();
            CopyOptions copyOpts = new CopyOptions
            {
                KeepMacros = true
            };
            sourceWithMacro.Copy(destWithMacro, copyOpts);
            destWithMacro.Save("Output_Scenario6_CopyWithMacros.xlsm", SaveFormat.Xlsm);

            Console.WriteLine("All scenarios executed successfully.");
        }

        private static void CreateSampleFile(string fileName)
        {
            if (!File.Exists(fileName))
            {
                Workbook wb = new Workbook();
                wb.Worksheets[0].Cells["A1"].PutValue($"Sample data for {fileName}");
                wb.Save(fileName, SaveFormat.Xlsx);
            }
        }

        private static void CreateMacroEnabledFile(string fileName)
        {
            if (!File.Exists(fileName))
            {
                Workbook wb = new Workbook();
                wb.Worksheets[0].Cells["A1"].PutValue("Macro enabled workbook");
                wb.Save(fileName, SaveFormat.Xlsm);
            }
        }
    }
}