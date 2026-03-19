using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMacroMapping
{
    public class ControlConfig
    {
        public string Type { get; set; }
        public int UpperLeftRow { get; set; }
        public int UpperLeftColumn { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public string MacroName { get; set; }
        public string LinkedCell { get; set; }
    }

    public class ConfigRoot
    {
        public List<ControlConfig> Controls { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and enable macros
            var workbook = new Workbook();
            workbook.Settings.EnableMacros = true;
            var sheet = workbook.Worksheets[0];

            // Load JSON configuration
            string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "controlsConfig.json");
            ConfigRoot config;

            if (File.Exists(jsonPath))
            {
                string jsonContent = File.ReadAllText(jsonPath);
                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                config = JsonSerializer.Deserialize<ConfigRoot>(jsonContent, options);
            }
            else
            {
                // Fallback sample configuration
                config = new ConfigRoot
                {
                    Controls = new List<ControlConfig>
                    {
                        new ControlConfig
                        {
                            Type = "Button",
                            UpperLeftRow = 1,
                            UpperLeftColumn = 1,
                            Width = 100,
                            Height = 30,
                            MacroName = "MyMacro",
                            LinkedCell = "A1"
                        }
                    }
                };
            }

            if (config?.Controls == null) return;

            foreach (var ctrl in config.Controls)
            {
                Shape shape = null;
                switch (ctrl.Type?.Trim().ToLower())
                {
                    case "button":
                        shape = sheet.Shapes.AddButton(ctrl.UpperLeftRow, ctrl.UpperLeftColumn, 0, 0, ctrl.Width, ctrl.Height);
                        break;
                    case "spinner":
                        shape = sheet.Shapes.AddSpinner(ctrl.UpperLeftRow, ctrl.UpperLeftColumn, 0, 0, ctrl.Width, ctrl.Height);
                        break;
                    case "listbox":
                        shape = sheet.Shapes.AddListBox(ctrl.UpperLeftRow, ctrl.UpperLeftColumn, 0, 0, ctrl.Width, ctrl.Height);
                        break;
                    default:
                        Console.WriteLine($"Unsupported control type: {ctrl.Type}");
                        continue;
                }

                if (!string.IsNullOrEmpty(ctrl.MacroName))
                    shape.MacroName = ctrl.MacroName;

                if (!string.IsNullOrEmpty(ctrl.LinkedCell))
                    shape.SetLinkedCell(ctrl.LinkedCell, false, false);

                shape.Name = $"{ctrl.Type}_{ctrl.UpperLeftRow}_{ctrl.UpperLeftColumn}";
            }

            // Save as macro‑enabled workbook
            workbook.Save("MappedControls.xlsm");
        }
    }
}