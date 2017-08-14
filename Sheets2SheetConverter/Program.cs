using Sheets2SheetConverter.Records;
using Newtonsoft.Json;
using System;
using System.IO;

namespace Sheets2SheetConverter
{
    class Program
    {
        static void Main (string[] args)
        {
            try
            { 
                string configFile = args[0];
            
                // Invalid config file
                if (!File.Exists (configFile))
                {
                    Console.WriteLine ("CONFIG FILE NOT FOUND");
                    return;
                }

                ConfigRecord config = JsonConvert.DeserializeObject<ConfigRecord>(File.ReadAllText (configFile));

                // Check input file
                if (!File.Exists (config.InputFile))
                {
                    Console.WriteLine ("INPUT FILE NOT FOUND");
                    return;
                }

                // Check output sheet name
                if (String.IsNullOrEmpty (config.OutputSheetName))
                {
                    Console.WriteLine ("OUTPUT SHEET NAME MUST BE FILLED OUT");
                    return;
                }

                // To avoid trash content inside an existing file
                if (File.Exists (config.OutputFile))
                {
                    File.Delete (config.OutputFile);
                }

                ExcelFileGenerator excelGen = new ExcelFileGenerator (config);
                
                // Try to convert file to a specific format
                if (!excelGen.ConvertFile ())
                {
                    Console.WriteLine (excelGen.Error);
                }
                else
                {
                    Console.WriteLine ("Concluído com sucesso!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine (ex.Message);
            }
        }
    }
}
