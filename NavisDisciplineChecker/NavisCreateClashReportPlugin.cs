using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.Plugins;

using DevExpress.Data.Async.Helpers;

using NavisDisciplineChecker.ClashModel;

namespace NavisDisciplineChecker {
    [Plugin("NavisCreateClashReportPlugin", "budaevaler",
        DisplayName = "Создать отчет коллизий", ToolTip = "Создает отчет коллизий")]
    public class NavisCreateClashReportPlugin : AddInPlugin {
        public override int Execute(params string[] parameters) {
            return CreateClashReportPlugin(parameters.FirstOrDefault() ?? DateTime.Now.ToString("yy.MM.dd"));
        }

        private int CreateClashReportPlugin(string currentDate) {
            var document = Application.ActiveDocument;

            var nwfFilePath = document.FileName;

            var logFileName = Path.ChangeExtension(nwfFilePath, ".log");
            var errorFileName = Path.ChangeExtension(nwfFilePath, ".error");

            var clashReportDirectoryPath =
                Path.Combine(Path.GetDirectoryName(document.FileName), "!Отчёты", currentDate);

            try {
                using(SimpleLogger logger = new SimpleLogger(logFileName)) {
                    DocumentClash clash = document.GetClash();

                    clash.TestsData.TestsRunAllTests();
                    logger.WriteLine("Запуск поиска коллизий.");

                    var clashes = clash.TestsData.Tests
                        .OfType<ClashTest>()
                        .Select(item => new {
                            TestName = item.DisplayName,
                            ClashResults = item.Children
                                .OfType<ClashResult>()
                                .Where(comment => comment.Status == ClashResultStatus.New
                                                  || comment.Status == ClashResultStatus.Active)
                                .ToList()
                        })
                        .Where(item => item.ClashResults.Count > 0);



                    var clashReports = clashes
                        .Select(item =>
                            new ClashReport() {
                                Name = item.TestName,
                                Clashes = item.ClashResults
                                    .Select(clashResult => new Clash(clashResult))
                                    .OrderBy(clashResult => clashResult.Name, new LogicalStringComparer())
                                    .ToList()
                            })
                        .ToList();

                    Directory.CreateDirectory(clashReportDirectoryPath);
                    logger.WriteLine($"Создание директории \"{clashReportDirectoryPath}\".");

                    foreach(ClashReport clashReport in clashReports) {
                        logger.WriteLine($"Создание отчета коллизии \"{clashReport.Name}\".");

                        using(StreamWriter stream =
                              new StreamWriter(Path.Combine(clashReportDirectoryPath, clashReport.Name + ".html"))) {
                            stream.WriteLine("<html>");
                            stream.WriteLine("<head>");
                            stream.WriteLine("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">");
                            stream.WriteLine($"<title>Отчет по коллизиям \"{clashReport.Name}\".</title>");
                            stream.WriteLine("<style type=\"text/css\">");
                            stream.WriteLine("table, th, td {");
                            stream.WriteLine("border: 1px solid black;");
                            stream.WriteLine("border-collapse: collapse;");
                            stream.WriteLine("}");
                            stream.WriteLine("</style>");
                            stream.WriteLine("</head>");


                            stream.WriteLine($"<caption>{clashReport.Name}</caption>");
                            stream.WriteLine("<table>");
                            stream.WriteLine("<tr>");

                            stream.WriteLine("<th>Id</th>");
                            stream.WriteLine("<th>Уровень</th>");
                            stream.WriteLine("<th>Категория</th>");
                            stream.WriteLine("<th>Имя типа</th>");
                            stream.WriteLine("<th>Имя файла</th>");

                            stream.WriteLine("<th>Id</th>");
                            stream.WriteLine("<th>Уровень</th>");
                            stream.WriteLine("<th>Категория</th>");
                            stream.WriteLine("<th>Имя типа</th>");
                            stream.WriteLine("<th>Имя файла</th>");

                            stream.WriteLine("</tr>");

                            foreach(Clash reportClash in clashReport.Clashes) {
                                stream.WriteLine("<tr>");
                                stream.WriteLine($"<td>{reportClash.ClashElement1.Id}</td>");
                                stream.WriteLine($"<td>{reportClash.ClashElement1.LevelName}</td>");
                                stream.WriteLine($"<td>{reportClash.ClashElement1.CategoryName}</td>");
                                stream.WriteLine($"<td>{reportClash.ClashElement1.TypeName}</td>");
                                stream.WriteLine($"<td>{reportClash.ClashElement1.SourceFileName}</td>");

                                stream.WriteLine($"<td>{reportClash.ClashElement2.Id}</td>");
                                stream.WriteLine($"<td>{reportClash.ClashElement2.LevelName}</td>");
                                stream.WriteLine($"<td>{reportClash.ClashElement2.CategoryName}</td>");
                                stream.WriteLine($"<td>{reportClash.ClashElement2.TypeName}</td>");
                                stream.WriteLine($"<td>{reportClash.ClashElement2.SourceFileName}</td>");
                                stream.WriteLine("</tr>");
                            }

                            stream.WriteLine("</html>");
                            stream.Flush();
                        }
                    }

                    return 0;
                }
            } catch(Exception ex) {
                File.WriteAllText(errorFileName, ex.ToString());
                return 1;
            }
        }
    }

    /// <summary>
    /// Умное сравнение строк.
    /// </summary>
    /// <remarks>Для сравнения используется метод WinApi <see cref="StrCmpLogicalW"/></remarks>
    public class LogicalStringComparer : IComparer<string> {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        private static extern int StrCmpLogicalW(string x, string y);

        /// <inheritdoc/>
        public int Compare(string x, string y) {
            return StrCmpLogicalW(x, y);
        }
    }
}