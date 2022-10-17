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
                    if(clash.TestsData.Tests.Count == 0) {
                        throw new InvalidOperationException(
                            $"У файла {Path.GetFileName(nwfFilePath)} не настроены правила коллизий.");
                    }

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

                        var htmlGrid =
                            new HtmlGrid(
                                    $"Отчет по коллизиям \"{clashReport.Name}\" + [{DateTime.Now.ToShortDateString()}].")
                                .CreateTitle($"Количество пересечений: {clashReport.Clashes.Count}")
                                .CreateColumns("№", "Id", "Уровень", "Категория", "Имя типа", "Имя файла");

                        int counter = 1;
                        foreach(Clash reportClash in clashReport.Clashes) {
                            htmlGrid.CreateRow(
                                counter++.ToString(),
                                reportClash.ClashElement1.Id,
                                reportClash.ClashElement1.LevelName,
                                reportClash.ClashElement1.CategoryName,
                                reportClash.ClashElement1.TypeName,
                                reportClash.ClashElement1.SourceFileName,
                                reportClash.ClashElement2.Id,
                                reportClash.ClashElement2.LevelName,
                                reportClash.ClashElement2.CategoryName,
                                reportClash.ClashElement2.TypeName,
                                reportClash.ClashElement2.SourceFileName);
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