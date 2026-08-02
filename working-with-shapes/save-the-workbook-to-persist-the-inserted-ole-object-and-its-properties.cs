// Title: Save Workbook with Embedded OLE Object and Preserve Its Properties – Aspose.Cells C#
// Description: Demonstrates how to add an OLE object to a worksheet, configure its ProgID, format, icon, and loading options, enable EmbedOoxmlAsOleObject, and save the workbook so the OLE object and all settings are retained.
// Keywords: Aspose.Cells C# OLE object | embed Excel file as OLE | save workbook with OLE icon | OoxmlSaveOptions EmbedOoxmlAsOleObject | persist OleObject properties | C# Excel OLE embedding | Aspose.Cells example
// Common Searches: how to embed an Excel file as an OLE object using Aspose.Cells | save OLE object with custom icon in .xlsx C# | preserve OLE object settings when exporting workbook | Aspose.Cells OoxmlSaveOptions embed OLE | C# code to add and persist OLE objects in Excel
// Developer Intent: Store the inserted OLE object and all its configured attributes by saving the workbook with appropriate options.
// Use Cases: Embedding a secondary workbook or document as an OLE object with a custom icon for distribution. | Setting OLE properties such as ProgID, FileFormatType, DisplayAsIcon, Label, AutoLoad, and AutoUpdate before saving. | Ensuring OLE objects are correctly written to an .xlsx file by using OoxmlSaveOptions.EmbedOoxmlAsOleObject.
// AI Prompts: Generate C# code that inserts a Word document as an OLE object, shows it as an icon, and saves the workbook with Aspose.Cells preserving the object. | Explain the impact of OoxmlSaveOptions.EmbedOoxmlAsOleObject on OLE storage in an Aspose.Cells workbook. | Provide a step‑by‑step guide to handle missing icon files when embedding OLE objects with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add an OLE object to a worksheet, configure its ProgID, format, icon, and loading options, enable EmbedOoxmlAsOleObject, and save the workbook so the OLE object and all settings are retained.
class SaveOleObjectDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Paths to the icon image and the file to embed
            string iconPath = "icon.png";
            string embedFilePath = "sample.xlsx";

            // Prepare icon image data (must be a valid image)
            byte[] iconData;
            if (File.Exists(iconPath))
            {
                iconData = File.ReadAllBytes(iconPath);
            }
            else
            {
                // Use a minimal 1x1 PNG as a placeholder (transparent pixel)
                const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
                iconData = Convert.FromBase64String(base64Png);
            }

            // Prepare embedded file data
            byte[] embedData;
            if (File.Exists(embedFilePath))
            {
                embedData = File.ReadAllBytes(embedFilePath);
            }
            else
            {
                // Create a minimal workbook to embed if the file is missing
                Workbook tempWb = new Workbook();
                using (MemoryStream ms = new MemoryStream())
                {
                    tempWb.Save(ms, SaveFormat.Xlsx);
                    embedData = ms.ToArray();
                }
            }

            // Add an OLE object using the icon image data
            int oleIndex = sheet.OleObjects.Add(2, 2, 100, 100, iconData);

            // Retrieve the OleObject instance
            OleObject ole = sheet.OleObjects[oleIndex];

            // Set the embedded file data and related properties
            ole.ObjectData = embedData;                     // actual embedded content
            ole.ProgID = "Excel.Sheet";                     // program identifier
            ole.FileFormatType = FileFormatType.Xlsx;       // format of the embedded file
            ole.DisplayAsIcon = true;                       // show as an icon
            ole.Label = "Embedded Sample.xlsx";             // icon label
            ole.AutoLoad = true;                            // load automatically when workbook opens
            ole.AutoUpdate = false;                         // do not auto‑update linked content

            // Create OOXML save options to embed OLE objects as OLE objects
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                EmbedOoxmlAsOleObject = true
            };

            // Save the workbook, persisting the OLE object and its properties
            workbook.Save("OleObjectPersisted.xlsx", saveOptions);
        }
        catch (Aspose.Cells.CellsException ex)
        {
            Console.WriteLine("Aspose.Cells error: " + ex.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
