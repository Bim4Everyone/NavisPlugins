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
            
        }
    }
}