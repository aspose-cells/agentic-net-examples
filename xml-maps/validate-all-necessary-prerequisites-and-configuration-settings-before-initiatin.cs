using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ValidateAndImportXmlMap
    {
        public static void Run()
        {
            // Paths for the source XML and the output Excel file
            string xmlFilePath = "data.xml";
            string outputPath = "output.xlsx";

            // ---------- Prerequisite validation ----------
            // Check that the XML file exists on disk
            if (!File.Exists(xmlFilePath))
            {
                Console.WriteLine($"Error: XML file not found at '{xmlFilePath}'.");
                return;
            }

            // ---------- Configuration settings ----------
            // Create XmlLoadOptions and enable XML mapping
            XmlLoadOptions loadOptions = new XmlLoadOptions
            {
                IsXmlMap = true,                     // Enable mapping of XML to Excel
                ContainsMultipleWorksheets = false,  // Single worksheet import (adjust as needed)
                ConvertNumericOrDate = true          // Convert numeric/date values automatically
            };

            // ---------- Load workbook with XML mapping ----------
            Workbook workbook;
            try
            {
                // Load the XML file using the configured options
                workbook = new Workbook(xmlFilePath, loadOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading XML with mapping: {ex.Message}");
                return;
            }

            // Verify that at least one XmlMap was created during load
            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XML maps were generated after loading. Ensure the XML has a valid schema.");
                return;
            }

            // ---------- Import/Update XML data ----------
            try
            {
                string targetSheetName = "Sheet1";

                // Ensure the target worksheet exists; add it if missing
                Worksheet sheet = workbook.Worksheets[targetSheetName];
                if (sheet == null)
                {
                    sheet = workbook.Worksheets.Add(targetSheetName);
                }

                // Import the XML data starting at cell A1 (row 0, column 0)
                workbook.ImportXml(xmlFilePath, targetSheetName, 0, 0);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during ImportXml: {ex.Message}");
                return;
            }

            // ---------- Save the workbook ----------
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateAndImportXmlMap.Run();
        }
    }
}