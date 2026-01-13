using Microsoft.Data.SqlClient;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using VertAccount.TechLabsLib.Activities.Properties;

namespace VertAccount.TechLabsLib.Activities
{
    [LocalizedDisplayName(nameof(Resources.CloseServer_DisplayName))]
    [LocalizedDescription(nameof(Resources.CloseServer_Description))]
    public class CloseServer : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.CloseServer_ConnectionString_DisplayName))]
        [LocalizedDescription(nameof(Resources.CloseServer_ConnectionString_Description))]
        [LocalizedCategory(nameof(Resources.Server_Category))]
        public InArgument<string> ConnectionString { get; set; }

        [LocalizedDisplayName(nameof(Resources.CloseServer_Result_DisplayName))]
        [LocalizedDescription(nameof(Resources.CloseServer_Result_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Result { get; set; }

        #endregion


        #region Constructors

        public CloseServer()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (ConnectionString == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(ConnectionString)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var connectionstring = ConnectionString.Get(context);

            ///////////////////////////
            // Add execution logic HERE
            using var conn = new SqlConnection(connectionstring);
            conn.Close();


            ///////////////////////////

            // Outputs
            return (ctx) => {
                Result.Set(ctx, true);
            };
        }

        #endregion
    }
}

