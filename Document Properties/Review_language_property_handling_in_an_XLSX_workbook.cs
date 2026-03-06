using System;
using Aspose.Cells;
using Aspose.Cells.Properties;
using Aspose.Cells.Drawing;

namespace AsposeCellsLanguageDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create a new workbook --------------------
            Workbook workbook = new Workbook(); // create

            // Set the built‑in document language property (e.g., en‑US)
            workbook.BuiltInDocumentProperties.Language = "en-US";

            // Set the workbook UI language code (e.g., German)
            workbook.Settings.LanguageCode = CountryCode.Germany;

            // Add a shape with text and set its language code (e.g., Japanese)
            Worksheet sheet = workbook.Worksheets[0];
            Shape shape = sheet.Shapes.AddTextBox(0, 0, 100, 100, 200, 50);
            shape.Text = "Sample text with language settings";
            shape.TextOptions.LanguageCode = CountryCode.Japan;

            // Save the workbook to disk
            workbook.Save("LanguageDemo.xlsx"); // save

            // -------------------- Load the workbook with specific LoadOptions --------------------
            LoadOptions loadOptions = new LoadOptions
            {
                // Assume the file was saved by a French user
                LanguageCode = CountryCode.France
            };
            Workbook loadedWorkbook = new Workbook("LanguageDemo.xlsx", loadOptions); // load

            // Output the language settings after loading
            Console.WriteLine("Built‑in Language Property: " + loadedWorkbook.BuiltInDocumentProperties.Language);
            Console.WriteLine("Workbook Settings LanguageCode: " + loadedWorkbook.Settings.LanguageCode);
            Console.WriteLine("LoadOptions LanguageCode (used for loading): " + loadOptions.LanguageCode);

            // Modify the built‑in language property to demonstrate write‑back
            loadedWorkbook.BuiltInDocumentProperties.Language = "fr-FR";

            // Save the modified workbook
            loadedWorkbook.Save("LanguageDemo_Loaded.xlsx"); // save
        }
    }
}