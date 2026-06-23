using System;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using Aspose.Cells;

namespace AsposeCellsThemeModification
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates a default workbook with default theme

            // 2. Save the workbook to a memory stream in XLSX format
            using (MemoryStream originalStream = new MemoryStream())
            {
                workbook.Save(originalStream, SaveFormat.Xlsx);
                originalStream.Position = 0; // reset for reading

                // 3. Open the XLSX package as a zip archive to access theme XML
                using (MemoryStream modifiedStream = new MemoryStream())
                {
                    // Copy original zip to a new stream that we can modify
                    originalStream.CopyTo(modifiedStream);
                    modifiedStream.Position = 0;

                    using (ZipArchive zip = new ZipArchive(modifiedStream, ZipArchiveMode.Update, true))
                    {
                        // Locate the theme part (usually xl/theme/theme1.xml)
                        ZipArchiveEntry themeEntry = zip.GetEntry("xl/theme/theme1.xml");
                        if (themeEntry != null)
                        {
                            // Load the theme XML
                            XDocument themeDoc;
                            using (Stream themeStream = themeEntry.Open())
                            {
                                themeDoc = XDocument.Load(themeStream);
                            }

                            // 4. Modify the <a:fontScheme> node (change major and minor fonts)
                            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

                            XElement fontScheme = themeDoc.Root.Element(a + "fontScheme");
                            if (fontScheme != null)
                            {
                                // Change major font latin typeface
                                XElement majorFont = fontScheme.Element(a + "majorFont");
                                if (majorFont != null)
                                {
                                    XElement latinMajor = majorFont.Element(a + "latin");
                                    if (latinMajor != null)
                                    {
                                        latinMajor.SetAttributeValue("typeface", "Calibri");
                                    }
                                }

                                // Change minor font latin typeface
                                XElement minorFont = fontScheme.Element(a + "minorFont");
                                if (minorFont != null)
                                {
                                    XElement latinMinor = minorFont.Element(a + "latin");
                                    if (latinMinor != null)
                                    {
                                        latinMinor.SetAttributeValue("typeface", "Arial");
                                    }
                                }
                            }

                            // Write the modified XML back into the zip entry
                            using (Stream themeWriteStream = themeEntry.Open())
                            {
                                // Truncate existing content
                                themeWriteStream.SetLength(0);
                                themeDoc.Save(themeWriteStream);
                            }
                        }
                    }

                    // 5. Load the modified workbook from the updated stream
                    modifiedStream.Position = 0;
                    Workbook modifiedWorkbook = new Workbook(modifiedStream);

                    // 6. Save the workbook with the updated theme
                    modifiedWorkbook.Save("ModifiedTheme.xlsx");
                }
            }

            Console.WriteLine("Workbook saved with updated font scheme in the theme.");
        }
    }
}