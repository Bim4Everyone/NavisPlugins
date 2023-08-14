using System.Collections.Generic;

namespace NavisDisciplineChecker.Tables {
    internal interface ITable {
        string Extension { get; }
        IReadOnlyList<Row> Rows { get; }
        IReadOnlyList<Column> Columns { get; }
        
        Row CreateRow();
        Column CreateColumn(string columnName);

        void SaveDocument(string filePath);
        void LoadDocument(string filePath);
    }
}