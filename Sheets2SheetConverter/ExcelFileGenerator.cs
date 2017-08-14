using BDCExcelManager;
using Sheets2SheetConverter.Records;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Sheets2SheetConverter
{
    public class SheetsData
    {
        /// <summary>
        /// Key = sheet name
        /// Value = lines data
        /// </summary>
        public Dictionary<string, List<string>> Data { get; set; }        

        /// <summary>
        /// Constructor
        /// </summary>
        public SheetsData ()
        {
            Data = new Dictionary<string, List<string>> ();
        }
    }

    public class ExcelFileGenerator
    {
        public ConfigRecord _config = null;
        private string      _error  = "";

        public string Error
        {
            get { return _error; }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="config">Config record</param>
        public ExcelFileGenerator (ConfigRecord config)
        {
            _config = config;
        }

        /// <summary>
        /// Convert file to a specific format
        /// </summary>
        /// <returns>True if no error occurred</returns>
        public bool ConvertFile ()
        {
            _error = null;
            try
            {
                // Load excel file into memory
                if (!LoadInputFile ())
                {
                    return false;
                }

                // Export to one single sheet file
                ExportOutputFile ();
                return true;
            }
            catch (Exception ex)
            {
                _error = ex.Message;
            }
            return false;
        }


        // File data loaded into memory
        private Dictionary<string, SheetsData> _data      = new Dictionary<string, SheetsData>();
        private Dictionary<string, string>     _headers   = new Dictionary<string, string> ();
        private string                         _keyHeader = "";

        /// <summary>
        /// Load Excel input file
        /// </summary>
        /// <returns>True if file was loaded</returns>
        private bool LoadInputFile ()
        {
            ExcelManager inputExcelManager = new ExcelManager (new FileInfo (_config.InputFile));
            foreach (SheetRecord sheetRec in _config.SheetRecords)
            {
                // Check if worksheet exists
                if (!inputExcelManager.WorksheetExists (sheetRec.SheetName))
                {
                    _error = "UNABLE TO FIND " + sheetRec.SheetName + " SHEET";
                    return false;
                }

                // We have to configure a limit to export
                if (sheetRec.MaxMatchs <= 0)
                {
                    _error = sheetRec.SheetName + " KEY MATCHES MUST BE BIGGER THAN 0";
                    return false;
                }

                // Open sheet and run over data
                Worksheet sheet = inputExcelManager.OpenWorksheet (sheetRec.SheetName);

                // Running over excel data
                for (int rowIndex = 1; rowIndex <= sheet.EPPlusSheet.Dimension.Rows; rowIndex++)
                {
                    int columnsCount   = sheet.EPPlusSheet.Dimension.Columns;
                    int keyColumnIndex = sheetRec.KeyIndex;

                    if (keyColumnIndex > columnsCount)
                    {
                        _error = "UNABLE TO FIND " + sheetRec.SheetName + " KEY INDEX " + keyColumnIndex;
                        return false;
                    }
                        
                    List<string>  sheetHeader = new List<string>();
                    StringBuilder lineData    = new StringBuilder ();
                    string        delimiter   = "";

                    foreach (int col in sheetRec.ColumnIndexes)
                    {
                        // Check if column exists
                        if (col > columnsCount)
                        {
                            _error = "UNABLE TO FIND " + sheetRec.SheetName + " INDEX " + col;
                            return false;
                        }
                            
                        string value = sheet.GetValue (rowIndex, col).ToString ();

                        // If we are at the first line
                        if (rowIndex == 1)
                        {
                            if (sheetRec.ContainsHeader)
                            {                                    
                                sheetHeader.Add (value);
                                continue;
                            }
                            else
                            {
                                sheetHeader.Add ("SHEET: " + sheetRec.SheetName + ", COLUMN: " + col);
                            }                            
                        }

                        lineData.Append (delimiter + value);
                        delimiter = sheetRec.ColumnsDelimiter;
                    }

                    // First time we are running we are checking key header
                    string key = sheet.GetValue (rowIndex, keyColumnIndex).ToString ();
                    if (String.IsNullOrEmpty (_keyHeader) && sheetRec.ContainsHeader)
                    {
                        _keyHeader = key;
                    }

                    // Load headers
                    if (sheetHeader.Count > 0 && !_headers.ContainsKey (sheetRec.SheetName))
                    {
                        _headers.Add (sheetRec.SheetName, String.Join (sheetRec.ColumnsDelimiter, sheetHeader));
                    }

                    // No match to write
                    if (lineData.Length == 0)
                    {
                        continue;
                    }

                    SheetsData   sheetsData   = null;
                    List<string> sheetContent = null;

                    if (_data.TryGetValue (key, out sheetsData))
                    {
                        if (sheetsData.Data.TryGetValue (sheetRec.SheetName, out sheetContent))
                        {
                            if (sheetRec.MaxMatchs > sheetContent.Count)
                            {
                                sheetContent.Add (lineData.ToString ());
                            }
                        }
                        else
                        {
                            sheetsData.Data.Add (sheetRec.SheetName, new List<string> () { lineData.ToString () });
                        }
                    }
                    else
                    {                            
                        sheetsData = new SheetsData ();                            
                        sheetsData.Data.Add (sheetRec.SheetName, new List<string> () { lineData.ToString () });
                        _data.Add (key, sheetsData);
                    }
                }
            }
            inputExcelManager.Dispose ();
            return true;
        }

        /// <summary>
        /// Export output file
        /// </summary>
        private void ExportOutputFile ()
        {
            ExcelManager outputExcelManager = new ExcelManager (new FileInfo (_config.OutputFile));
            Worksheet    outSheet           = outputExcelManager.OpenOrCreateWorksheet (_config.OutputSheetName);

            // Key header
            outSheet.Write (1, 1, _keyHeader);
            outSheet.SetBold (1, 1);

            // Run over docs
            int outRow = 1;
            foreach (KeyValuePair<string, SheetsData> doc in _data)
            {
                int outCol = 2;
                outRow++;

                // Key at the first column always (ex.: 'Documento')
                outSheet.Write (outRow, 1, doc.Key);

                // Run over sheets to keep the sorting order
                foreach (SheetRecord sheetRec in _config.SheetRecords)
                {
                    List<string> docData   = null;
                    string       content   = "";
                    int          remaining = sheetRec.MaxMatchs;

                    if (doc.Value.Data.TryGetValue (sheetRec.SheetName, out docData))
                    {
                        content   = String.Join (sheetRec.LinesDelimiter, docData);
                        remaining = sheetRec.MaxMatchs - sheetRec.MaxMatchs;
                    }

                    // Just fill with empty delimiters
                    // Since we try to avoid empty lines like ';;;;;' we only consider tab delimiter
                    // tab delimiter is used to separate columns
                    if (remaining != sheetRec.MaxMatchs && sheetRec.LinesDelimiter.Equals ("\t", StringComparison.OrdinalIgnoreCase))
                    {
                        for (int i = 0; i < remaining; i++)
                        {
                            content += sheetRec.LinesDelimiter;
                        }
                    }

                    // Write in output file
                    int headerCol = outCol;
                    foreach (string c in content.Split ('\t'))
                    {
                        outSheet.Write (outRow, outCol, c);
                        outCol++;
                    }

                    if (_headers.ContainsKey (sheetRec.SheetName))
                    {
                        string currentHeaders = String.Join (sheetRec.ColumnsDelimiter, _headers[sheetRec.SheetName]);

                        // Write headers in output file
                        int i = 0;
                        do
                        {
                            i++;
                            foreach (string c in currentHeaders.Split ('\t'))
                            {
                                outSheet.Write (1, headerCol, c);
                                outSheet.SetBold (1, headerCol);
                                headerCol++;
                            }
                         
                        } while (i < remaining);
                    }
                }
            }
            // Close and save files
            outputExcelManager.Save ();
            outputExcelManager.Dispose ();
        }
    }
}
