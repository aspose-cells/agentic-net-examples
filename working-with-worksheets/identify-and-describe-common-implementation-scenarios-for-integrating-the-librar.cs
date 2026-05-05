using System;
using System.IO;
using System.IO.MemoryMappedFiles;
using Aspose.Cells;
using Aspose.Cells.WebExtensions;
using Aspose.Cells.Drawing;
using Aspose.Cells.Vba;

namespace AsposeCellsIntegrationScenarios
{
    class Program
    {
        static void Main(string[] args)
        {
            // 1. Licensing the Aspose.Cells component (required for all scenarios)
            License license = new License();
            string licensePath = "Aspose.Cells.NET.lic";
            if (File.Exists(licensePath))
            {
                license.SetLicense(licensePath);
                Console.WriteLine("License applied from file.");
            }
            else
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    // license.SetLicense(ms);
                    Console.WriteLine("License stream would be applied here.");
                }
            }

            // 2. Setting global helper properties (affects all workbook operations)
            CellsHelper.IsCloudPlatform = true;
            CellsHelper.StartupPath = "/cloud_storage/aspose/cells";
            CellsHelper.LibraryPath = "/cloud_storage/aspose/cells/libs";
            Console.WriteLine($"IsCloudPlatform: {CellsHelper.IsCloudPlatform}");
            Console.WriteLine($"StartupPath: {CellsHelper.StartupPath}");
            Console.WriteLine($"LibraryPath: {CellsHelper.LibraryPath}");

            // 3. Creating a workbook, adding data, and saving (basic usage)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
            sheet.Cells["A2"].Formula = "=SUM(1,2,3)";
            workbook.Save("BasicWorkbook.xlsx");
            Console.WriteLine("Basic workbook saved.");

            // 4. Using Metered licensing (for metered usage scenarios)
            try
            {
                Metered metered = new Metered();
                string publicKey = "YourPublicKey";
                string privateKey = "YourPrivateKey";
                metered.SetMeteredKey(publicKey, privateKey);
                Console.WriteLine($"Metered product name: {metered.GetProductName()}");
                Console.WriteLine($"Is metered licensed: {Metered.IsMeteredLicensed()}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Metered licensing not applied: {ex.Message}");
            }

            // 5. Registering and using an Excel add‑in function (UDF)
            string addInPath = Path.Combine(Environment.CurrentDirectory, "MyAddIn.xlam");
            if (File.Exists(addInPath))
            {
                int addInId = workbook.Worksheets.RegisterAddInFunction(addInPath, "MY_UDF", false);
                sheet.Cells["B1"].Formula = "=MY_UDF()";
                Console.WriteLine($"Add‑in function registered with ID {addInId}.");
            }
            else
            {
                Console.WriteLine("Add‑in file not found; skipping UDF registration.");
            }

            // 6. Embedding a WebExtension (e.g., a YouTube video) into a worksheet
            WebExtensionCollection webExts = workbook.Worksheets.WebExtensions;
            int webExtIndex = webExts.Add();
            WebExtension webExt = webExts[webExtIndex];
            webExt.Reference.Id = "youtube";
            webExt.Reference.StoreName = "YouTube";
            webExt.Properties.Add("videoUrl", "https://www.youtube.com/watch?v=dQw4w9WgXcQ");

            ShapeCollection shapes = sheet.Shapes;
            shapes.AddShape(MsoDrawingType.WebExtension, 5, 1, 0, 0, 200, 300);
            WebExtensionShape webShape = (WebExtensionShape)shapes[shapes.Count - 1];
            webShape.WebExtension = webExt;
            Console.WriteLine("WebExtension added to worksheet.");

            // Save final workbook with all features.
            workbook.Save("IntegratedFeaturesWorkbook.xlsx");
            Console.WriteLine("Final workbook with integrated features saved.");
        }
    }
}