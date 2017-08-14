using System.Collections.Generic;

namespace Sheets2SheetConverter.Records
{
    public class ConfigRecord
    {
        public string            InputFile       { get; set; }
        public string            OutputFile      { get; set; }
        public string            OutputSheetName { get; set; }
        public List<SheetRecord> SheetRecords    { get; set; }
    }

    public class SheetRecord
    {
        public int       KeyIndex         { get; set; }
        public int       MaxMatchs        { get; set; }        
        public bool      ContainsHeader   { get; set; }        
        public string    SheetName        { get; set; }
        public string    ColumnsDelimiter { get; set; }
        public string    LinesDelimiter   { get; set; }        
        public List<int> ColumnIndexes    { get; set; }
    }
}
