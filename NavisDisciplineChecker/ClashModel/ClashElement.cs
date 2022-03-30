using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Navisworks.Api;

namespace NavisDisciplineChecker.ClashModel {
    public class ClashElement {
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
