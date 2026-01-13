using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using VertAccount.TechLabsLib.Activities.Design.Designers;
using VertAccount.TechLabsLib.Activities.Design.Properties;

namespace VertAccount.TechLabsLib.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(Server), categoryAttribute);
            builder.AddCustomAttributes(typeof(Server), new DesignerAttribute(typeof(ServerDesigner)));
            builder.AddCustomAttributes(typeof(Server), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(GetData), categoryAttribute);
            builder.AddCustomAttributes(typeof(GetData), new DesignerAttribute(typeof(GetDataDesigner)));
            builder.AddCustomAttributes(typeof(GetData), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(DataOuput), categoryAttribute);
            builder.AddCustomAttributes(typeof(DataOuput), new DesignerAttribute(typeof(DataOuputDesigner)));
            builder.AddCustomAttributes(typeof(DataOuput), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(VerifyVersion), categoryAttribute);
            builder.AddCustomAttributes(typeof(VerifyVersion), new DesignerAttribute(typeof(VerifyVersionDesigner)));
            builder.AddCustomAttributes(typeof(VerifyVersion), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(PDFMerger), categoryAttribute);
            builder.AddCustomAttributes(typeof(PDFMerger), new DesignerAttribute(typeof(PDFMergerDesigner)));
            builder.AddCustomAttributes(typeof(PDFMerger), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(CloseServer), categoryAttribute);
            builder.AddCustomAttributes(typeof(CloseServer), new DesignerAttribute(typeof(CloseServerDesigner)));
            builder.AddCustomAttributes(typeof(CloseServer), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
