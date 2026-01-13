using Microsoft.Data.SqlClient;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using VertAccount.TechLabsLib.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;


namespace VertAccount.TechLabsLib.Activities
{
    [LocalizedDisplayName(nameof(Resources.Server_DisplayName))]
    [LocalizedDescription(nameof(Resources.Server_Description))]
    public class Server : ContinuableAsyncCodeActivity
    {
        #region Properties
        public enum authType
        {
            SQLServer,
            Windows
        }
        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.Server_ServerAuthType_DisplayName))]
        [LocalizedDescription(nameof(Resources.Server_ServerAuthType_Description))]
        [LocalizedCategory(nameof(Resources.Server_Category))]
        public authType ServerAuthType { get; set; }

        [LocalizedDisplayName(nameof(Resources.Server_ServerName_DisplayName))]
        [LocalizedDescription(nameof(Resources.Server_ServerName_Description))]
        [LocalizedCategory(nameof(Resources.Server_Category))]
        public InArgument<string> ServerName { get; set; }

        [LocalizedDisplayName(nameof(Resources.Server_Database_DisplayName))]
        [LocalizedDescription(nameof(Resources.Server_Database_Description))]
        [LocalizedCategory(nameof(Resources.Server_Category))]
        public InArgument<string> Database { get; set; }

        [LocalizedDisplayName(nameof(Resources.Server_UserName_DisplayName))]
        [LocalizedDescription(nameof(Resources.Server_UserName_Description))]
        [LocalizedCategory(nameof(Resources.Authentication_Category))]
        public InArgument<string> UserName { get; set; }

        [LocalizedDisplayName(nameof(Resources.Server_Password_DisplayName))]
        [LocalizedDescription(nameof(Resources.Server_Password_Description))]
        [LocalizedCategory(nameof(Resources.Authentication_Category))]
        public InArgument<string> Password { get; set; }

        [LocalizedDisplayName(nameof(Resources.Server_ConnectionString_DisplayName))]
        [LocalizedDescription(nameof(Resources.Server_ConnectionString_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> ConnectionString { get; set; }

        [LocalizedDisplayName(nameof(Resources.Server_Connected_DisplayName))]
        [LocalizedDescription(nameof(Resources.Server_Connected_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<bool> Connected { get; set; }

        #endregion


        #region Constructors

        public Server()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (ServerName == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(ServerName)));
            if (Database == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Database)));
            if (UserName == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(UserName)));
            if (Password == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Password)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            //AppContext.SetSwitch("Microsoft.Data.SqlClient.UseManagedNetworkingOnWindows", true);
            // Inputs
            var servername = ServerName.Get(context);
            var database = Database.Get(context);
            var username = UserName.Get(context);
            var password = Password.Get(context);
            bool isConnected;
            string connectionString;

            ///////////////////////////
            // Add execution logic HERE

            if (ServerAuthType == authType.SQLServer)
            {
                //SQL Server authentication
                connectionString = "Server=" + servername + ";Database=" + database + ";User Id=" + username + ";Password=" + password + ";TrustServerCertificate=True;";
            }

            else
            {
                //Windows authentication
                connectionString = "Server=" + servername + ";Database=" + database + ";Integrated Security=True;TrustServerCertificate=True;";
            }

            try
            {
                using var conn = new SqlConnection(connectionString);
                conn.Open();
                isConnected = true;
                Console.WriteLine("Test Connect to DB Succeed!");
                conn.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.Message);
                Console.WriteLine("INNER: " + ex.InnerException?.Message);
                isConnected = false;
            }

            ///////////////////////////

            // Outputs
            return (ctx) => {
                ConnectionString.Set(ctx, connectionString);
                Connected.Set(ctx, isConnected);
            };
        }

        #endregion
    }
}

