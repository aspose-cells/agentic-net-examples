using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPrerequisiteCheck
{
    class Program
    {
        static void Main()
        {
            // 1. Detect cloud platform
            string cloudEnv = Environment.GetEnvironmentVariable("ASPOSE_CLOUD_PLATFORM");
            bool isCloud = !string.IsNullOrEmpty(cloudEnv) && cloudEnv.Equals("true", StringComparison.OrdinalIgnoreCase);
            CellsHelper.IsCloudPlatform = isCloud;
            Console.WriteLine($"CellsHelper.IsCloudPlatform set to: {CellsHelper.IsCloudPlatform}");

            // 2. Load Aspose.Cells license
            string licenseEnv = Environment.GetEnvironmentVariable("ASPOSE_CELLS_LICENSE");
            string licensePath = !string.IsNullOrEmpty(licenseEnv)
                ? licenseEnv
                : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Aspose.Cells.lic");

            if (File.Exists(licensePath))
            {
                try
                {
                    var license = new License();
                    license.SetLicense(licensePath);
                    Console.WriteLine("Aspose.Cells license loaded successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to set license: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"License file not found at: {licensePath}");
                Console.WriteLine("Proceeding without a license (evaluation mode).");
            }

            // 3. Verify Aspose.Cells library version
            try
            {
                string version = CellsHelper.GetVersion();
                Console.WriteLine($"Aspose.Cells version: {version}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unable to retrieve Aspose.Cells version: {ex.Message}");
            }

            // 4. Confirm required dependencies (fonts) are available via Aspose.Cells
            try
            {
                var wb = new Workbook();
                string defaultFont = wb.DefaultStyle.Font.Name;
                Console.WriteLine($"Default workbook font: {defaultFont}");
                Console.WriteLine("Font dependency is available.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Font dependency check failed: {ex.Message}");
            }

            // 5. Summary
            Console.WriteLine("\nPrerequisite verification completed.");
        }
    }
}