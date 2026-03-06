using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtection
{
    // Model representing the JSON configuration
    public class ProtectionConfig
    {
        public string InputPath { get; set; }
        public string OutputPath { get; set; }
        public string Password { get; set; }
        public bool LockForViewing { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Example JSON configuration (could be read from a file or other source)
            string json = @"
            {
                ""InputPath"": ""sample.xlsm"",
                ""OutputPath"": ""sample_protected.xlsm"",
                ""Password"": ""mySecretPwd"",
                ""LockForViewing"": true
            }";

            // Deserialize JSON into the configuration object
            ProtectionConfig config = JsonSerializer.Deserialize<ProtectionConfig>(json);

            // Load the workbook that contains a VBA project
            Workbook workbook = new Workbook(config.InputPath);

            // Protect the VBA project with the specified password and lock setting
            // isLockedForViewing = true means the project cannot be viewed without the password
            workbook.VbaProject.Protect(config.LockForViewing, config.Password);

            // Save the workbook as a macro‑enabled file (Xlsm) to retain the VBA project
            workbook.Save(config.OutputPath, SaveFormat.Xlsm);
        }
    }
}