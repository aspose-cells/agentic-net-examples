using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

public class XlsxUsageScenarios
{
    // 1. Create a new workbook and save it as XLSX
    public static void CreateAndSaveXlsx()
    {
        // Create a new workbook (default format is Xlsx)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello XLSX");

        // Save the workbook using the SaveFormat enumeration
        workbook.Save("CreateExample.xlsx", SaveFormat.Xlsx);
    }

    // 2. Load an existing XLSX file, modify it, and save to a new file
    public static void LoadFromFileAndSave()
    {
        // Load workbook from file path
        Workbook workbook = new Workbook("CreateExample.xlsx");

        // Modify the workbook
        workbook.Worksheets[0].Cells["B1"].PutValue(DateTime.Now);

        // Save the modified workbook as a new XLSX file
        workbook.Save("ModifiedExample.xlsx", SaveFormat.Xlsx);
    }

    // 3. Load an XLSX file from a stream, edit, and save back to disk
    public static void LoadFromStreamAndSave()
    {
        // Open the source file as a FileStream
        using (FileStream fileStream = new FileStream("CreateExample.xlsx", FileMode.Open, FileAccess.Read))
        {
            // Copy to a MemoryStream for Aspose.Cells consumption
            using (MemoryStream memoryStream = new MemoryStream())
            {
                fileStream.CopyTo(memoryStream);
                memoryStream.Position = 0; // Reset position for reading

                // Load workbook from the memory stream
                Workbook workbook = new Workbook(memoryStream);

                // Add some data
                workbook.Worksheets[0].Cells["C1"].PutValue(123);

                // Save to another MemoryStream
                using (MemoryStream outStream = new MemoryStream())
                {
                    workbook.Save(outStream, SaveFormat.Xlsx);
                    // Write the resulting bytes to a file
                    File.WriteAllBytes("StreamSaved.xlsx", outStream.ToArray());
                }
            }
        }
    }

    // 4. Convert an XLSX file to PDF using the ConversionUtility
    public static void ConvertXlsxToPdf()
    {
        // Convert directly without loading the workbook into memory
        ConversionUtility.Convert("CreateExample.xlsx", "Converted.pdf");
    }

    // 5. Detect the file format of an existing file
    public static void DetectFileFormat()
    {
        // Use FileFormatUtil to detect format information
        FileFormatInfo info = FileFormatUtil.DetectFileFormat("CreateExample.xlsx");
        Console.WriteLine($"Detected format: {info.FileFormatType}");
    }

    // 6. Save an XLSX file with strict OOXML compliance
    public static void SaveWithOoxmlCompliance()
    {
        Workbook workbook = new Workbook();

        // Set compliance level to ISO/IEC 29500:2008 Strict
        workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        workbook.Worksheets[0].Cells["A1"].PutValue("Strict OOXML");
        workbook.Save("StrictCompliance.xlsx", SaveFormat.Xlsx);
    }

    // Entry point to run all scenarios
    public static void Main()
    {
        CreateAndSaveXlsx();
        LoadFromFileAndSave();
        LoadFromStreamAndSave();
        ConvertXlsxToPdf();
        DetectFileFormat();
        SaveWithOoxmlCompliance();

        Console.WriteLine("All XLSX usage scenarios have been executed.");
    }
}