using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Navisworks.Api;

namespace NavisDisciplineChecker.ClashModel {
    public class ClashElement {
        public ClashElement(ModelItem modelItem) {
            if(modelItem == null) {
                throw new ArgumentNullException(nameof(modelItem));
            }

            Id = ToDisplayString(modelItem, "LcRevitId", "LcOaNat64AttributeValue");
            CategoryName = ToDisplayString(modelItem, "LcRevitData_Element", "LcRevitPropertyElementCategory");
            TypeName = ToDisplayString(modelItem, "LcRevitData_Element", "LcRevitPropertyElementType");
            LevelName = ToDisplayString(modelItem, "ReferenceLevel", "Name");
            SourceFileName = ToDisplayString(modelItem, "LcOaNode", "LcOaNodeSourceFile");
        }

        public string Id { get; set; }
        public string CategoryName { get; set; }
        public string TypeName { get; set; }
        public string LevelName { get; set; }
        public string SourceFileName { get; set; }

        private static string ToDisplayString(ModelItem modelItem, string categoryName, string propertyName) {
            return modelItem?.Ancestors
                ?.First?.PropertyCategories
                ?.FindCategoryByName(categoryName)
                ?.Properties?.FindPropertyByName(propertyName)?.Value?.ToDisplayString();
        }
    }
}
