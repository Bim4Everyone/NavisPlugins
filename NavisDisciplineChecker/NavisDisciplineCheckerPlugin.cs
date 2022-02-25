using System;
using Autodesk.Navisworks.Api.Plugins;

namespace NavisDisciplineChecker {

    [Plugin("NavisDisciplineChecker", "budaevaler", DisplayName = "Проверка nwf", ToolTip = "Проверка nwf по разделам")]
    public class NavisDisciplineCheckerPlugin : AddInPlugin {
        public override int Execute(params string[] parameters) {
            return 0;
        }
    }
}
