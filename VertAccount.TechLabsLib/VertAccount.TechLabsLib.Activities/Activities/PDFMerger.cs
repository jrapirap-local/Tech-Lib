using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using VertAccount.TechLabsLib.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace VertAccount.TechLabsLib.Activities
{
    [LocalizedDisplayName(nameof(Resources.PDFMerger_DisplayName))]
    [LocalizedDescription(nameof(Resources.PDFMerger_Description))]
    public class PDFMerger : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.PDFMerger_Path_DisplayName))]
        [LocalizedDescription(nameof(Resources.PDFMerger_Path_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Path { get; set; }

        [LocalizedDisplayName(nameof(Resources.PDFMerger_PDFFiles_DisplayName))]
        [LocalizedDescription(nameof(Resources.PDFMerger_PDFFiles_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<List<string>> PDFFiles { get; set; }

        [LocalizedDisplayName(nameof(Resources.PDFMerger_PDFOutputName_DisplayName))]
        [LocalizedDescription(nameof(Resources.PDFMerger_PDFOutputName_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> PDFOutputName { get; set; }

        #endregion


        #region Constructors

        public PDFMerger()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Path == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Path)));
            if (PDFFiles == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(PDFFiles)));
            if (PDFOutputName == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(PDFOutputName)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var path = Path.Get(context);
            var pdffiles = PDFFiles.Get(context);
            var pdfoutputname = PDFOutputName.Get(context);
            var outPath = path + "\\" + pdfoutputname;

            ///////////////////////////
            // Add execution logic HERE

            using (PdfDocument outputDocument = new PdfDocument())
            {
                foreach (string file in pdffiles)
                {
                    using (PdfDocument inputDocument = PdfReader.Open(file, PdfDocumentOpenMode.Import))
                    {
                        for (int idx = 0; idx < inputDocument.PageCount; idx++)
                        {
                            PdfPage page = inputDocument.Pages[idx];
                            outputDocument.AddPage(page);
                        }
                    }
                }

                outputDocument.Save(outPath);
            }
            ///////////////////////////

            // Outputs
            return (ctx) => {
            };
        }

        #endregion
    }
}

