using System.Collections.Generic;
using System.Linq;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.Plugins;

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

    internal class ClashReport {
        public string Name { get; set; }
        public List<Clash> Clashes { get; set; } = new List<Clash>();
    }

    internal class Clash {
        public Clash(ClashResult clashResult) {
            Name = clashResult.DisplayName;
            ClashElement1 = new ClashElement(clashResult.Item1);
            ClashElement2 = new ClashElement(clashResult.Item2);
        }

        public string Name { get; set; }
        public ClashElement ClashElement1 { get; set; }
        public ClashElement ClashElement2 { get; set; }
    }

    internal class ClashElement {
        public ClashElement(ModelItem modelItem) {
            Id = ToDisplayString(modelItem, "LcRevitId", "LcOaNat64AttributeValue");
            Name = ToDisplayString(modelItem, "LcOaNode", "LcOaSceneBaseUserName");
            TypeName = ToDisplayString(modelItem, "LcOaNode", "LcOaSceneBaseClassUserName");
            LayerName = ToDisplayString(modelItem, "LcOaNode", "LcOaNodeLayer");
            SourceFileName = ToDisplayString(modelItem, "LcOaNode", "LcOaNodeSourceFile");
        }

        public string Id { get; set; }
        public string Name { get; set; }
        public string TypeName { get; set; }
        public string LayerName { get; set; }
        public string SourceFileName { get; set; }

        private static string ToDisplayString(ModelItem modelItem, string categoryName, string propertyName) {
            return modelItem.Ancestors
                .First.PropertyCategories
                .FindCategoryByName(categoryName)
                ?.Properties.FindPropertyByName(propertyName).Value?.ToDisplayString();
        }
    }
}