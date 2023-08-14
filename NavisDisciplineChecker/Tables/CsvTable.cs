using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NavisDisciplineChecker.Tables {
    internal sealed class CsvTable : ITable {
        private readonly List<Row> _rows = new List<Row>();
        private readonly List<Column> _columns = new List<Column>();

        public char Separator => ';';
        public string Extension => ".csv";
        public Encoding Encoding => Encoding.UTF8;

        public IReadOnlyList<Row> Rows => _rows;
        public IReadOnlyList<Column> Columns => _columns;

        public Row CreateRow() {
            var newRow = new Row(this);
            _rows.Add(newRow);

            foreach(var column in Columns) {
                newRow.InsertColumn(column);
            }

            return newRow;
        }

        public Column CreateColumn(string columnName) {
            var newColumn = new Column(this, Columns.Count, columnName);
            _columns.Add(newColumn);

            foreach(var row in Rows) {
                row.InsertColumn(newColumn);
            }

            return newColumn;
        }

        public void SaveDocument(string filePath) {
            CheckExtension(filePath);

            using(StreamWriter stream = new StreamWriter(filePath, false, Encoding)) {
                stream.WriteLine(
                    string.Join(
                        Separator.ToString(),
                        Columns
                            .OrderBy(item => item.Index)
                            .Select(item => item.ColumnName)));

                foreach(Row row in Rows) {
                    stream.WriteLine(string.Join(Separator.ToString(), row));
                }
            }
        }

        public void LoadDocument(string filePath) {
            CheckExtension(filePath);
            using(StreamReader stream = new StreamReader(filePath, Encoding)) {
                foreach(string columnName in ReadLine(stream)) {
                    CreateColumn(columnName);
                }

                while(!stream.EndOfStream) {
                    var row = CreateRow();

                    int index = 0;
                    foreach(string rowValue in ReadLine(stream)) {
                        row[index++] = rowValue;
                    }
                }
            }
        }

        private void CheckExtension(string filePath) {
            if(!filePath.EndsWith(Extension)) {
                throw new ArgumentException("Не поддерживаемое расширение.");
            }
        }

        private IEnumerable<string> ReadLine(StreamReader stream) {
            return stream.ReadLine()?
                .Split(Separator) ?? Enumerable.Empty<string>();
        }
    }

    internal sealed class Row : IEnumerable<string> {
        private readonly CsvTable _csvTable;
        private readonly List<string> _row;

        public Row(CsvTable csvTable) {
            _csvTable = csvTable;
            _row = new List<string>(csvTable.Columns.Count);
        }

        public string this[int index] {
            get => _row[index];
            set => _row[index] = value;
        }

        public string this[Column column] {
            get => this[column.Index];
            set => this[column.Index] = value;
        }

        public void AddColumn(string value = default) {
            _row.Add(value);
        }

        public void InsertColumn(Column column, string value = default) {
            _row.Insert(column.Index, value);
        }

        public IEnumerator<string> GetEnumerator() {
            return _row.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator() {
            return GetEnumerator();
        }
    }

    internal sealed class Column {
        private readonly CsvTable _csvTable;

        public Column(CsvTable csvTable, int index, string columnName) {
            _csvTable = csvTable;

            Index = index;
            ColumnName = columnName;
        }

        public int Index { get; }
        public string ColumnName { get; }
    }
}