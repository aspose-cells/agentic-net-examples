// Title: Set workbook culture to format optional property values during export with Aspose.Cells for .NET
// Common Searches: Aspose.Cells set workbook cultureinfo to French before saving | C# export Excel with locale-specific date and number formatting using Aspose.Cells | How to apply culture settings to optional properties in Aspose.Cells workbook | Configure workbook.Settings.CultureInfo for Excel export in .NET | Change workbook culture for culture-aware formatting in Aspose.Cells
// Tags: workbook.Settings.CultureInfo Aspose.Cells | locale specific formatting workbook export | culture aware number date formatting Aspose.Cells | set workbook culture for Excel save .NET | optional property value culture formatting Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// Demonstrates setting Workbook.Settings.CultureInfo (e.g., "fr-FR") so that dates and numbers are formatted according to the specified locale when the workbook is saved.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // create

        // Set the culture that will be used for formatting optional property values during export
        // For example, use French (France) culture
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Optional: add some data to demonstrate culture-specific formatting
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(DateTime.Now);          // Date will be formatted according to the set culture
        sheet.Cells["A2"].PutValue(12345.67);              // Number will be formatted according to the set culture

        // Save the workbook to a file (export)
        workbook.Save("ExportedWorkbook.xlsx"); // save
    }
}
