using System;
using System.Linq;
using System.Text;

namespace NavisDisciplineChecker {
    internal class HtmlGrid {
        private readonly StringBuilder _builder = new StringBuilder();

        public HtmlGrid(string titleName) {
            _builder.Append("<!DOCTYPE HTML>");
            _builder.Append("<html>");
            _builder.Append("<head>");
            _builder.Append("<meta charset=\"utf-8\">");
            _builder.Append("<title>");
            _builder.Append(titleName);
            _builder.Append("</title>");
            _builder.Append("<style>");
            _builder.Append(@"table, th, td {
  border: 1px solid silver;
  border-collapse: collapse;
  margin: 10px;
  padding: 10px;
}
tr:nth-child(even) {
  background-color: #D6EEEE;
}

caption {
    text-align: left;
    font-size:200%;
    margin: 10px;
}");
            _builder.Append("</style>");
            _builder.Append("</head>");
            _builder.Append("<body>");
            _builder.Append("<table style=\"width:100%\">");
        }

        public HtmlGrid CreateTitle(string tableTitle) {
            _builder.Append("<caption>");
            _builder.Append("<b>");
            _builder.Append(tableTitle);
            _builder.Append("</b>");
            _builder.Append("</caption>");

            return this;
        }

        public HtmlGrid CreateColumns(params string[] columnNames) {
            _builder.Append("<tr>");
            _builder.Append(string.Join(Environment.NewLine, columnNames.Select(item => $"<th>{item}</th>")));
            _builder.Append("</tr>");

            return this;
        }

        public HtmlGrid CreateRow(params string[] dataRow) {
            _builder.Append("<tr>");
            _builder.Append(string.Join(Environment.NewLine, dataRow.Select(item => $"<td>{item}</td>")));
            _builder.Append("</tr>");
            return this;
        }

        public override string ToString() {
            _builder.Append("</table>");
            _builder.Append("</body>");
            _builder.Append("</html>");
            return _builder.ToString();
        }
    }
}