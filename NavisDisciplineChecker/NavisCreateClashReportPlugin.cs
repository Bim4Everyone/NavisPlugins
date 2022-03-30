using System.Linq;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.Plugins;

using NavisDisciplineChecker.ClashModel;

namespace NavisDisciplineChecker {
    [Plugin("NavisCreateClashReportPlugin", "budaevaler",
        DisplayName = "Создать отчет коллизий", ToolTip = "Создает отчет коллизий")]
    public class NavisCreateClashReportPlugin : AddInPlugin {
        public override int Execute(params string[] parameters) {
            CreateClashReportPlugin();
            return 0;
        }

        private void CreateClashReportPlugin() {
            var document = Application.ActiveDocument;

            DocumentClash clash = document.GetClash();
            clash.TestsData.TestsRunAllTests();

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
                            .ToList()
                    })
                .ToList();
        }
    }
}