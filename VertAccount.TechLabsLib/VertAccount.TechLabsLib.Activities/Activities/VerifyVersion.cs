using Microsoft.Data.SqlClient;
using System;
using System.Activities;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using VertAccount.TechLabsLib.Activities.Properties;

namespace VertAccount.TechLabsLib.Activities
{
    [LocalizedDisplayName(nameof(Resources.VerifyVersion_DisplayName))]
    [LocalizedDescription(nameof(Resources.VerifyVersion_Description))]
    public class VerifyVersion : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.VerifyVersion_ConnectionString_DisplayName))]
        [LocalizedDescription(nameof(Resources.VerifyVersion_ConnectionString_Description))]
        [LocalizedCategory(nameof(Resources.Server_Category))]
        public InArgument<string> ConnectionString { get; set; }

        [LocalizedDisplayName(nameof(Resources.VerifyVersion_BotName_DisplayName))]
        [LocalizedDescription(nameof(Resources.VerifyVersion_BotName_Description))]
        [LocalizedCategory(nameof(Resources.VABot_Category))]
        public InArgument<string> BotName { get; set; }

        [LocalizedDisplayName(nameof(Resources.VerifyVersion_BotVersion_DisplayName))]
        [LocalizedDescription(nameof(Resources.VerifyVersion_BotVersion_Description))]
        [LocalizedCategory(nameof(Resources.VABot_Category))]
        public InArgument<string> BotVersion { get; set; }

        [LocalizedDisplayName(nameof(Resources.VerifyVersion_ResultText_DisplayName))]
        [LocalizedDescription(nameof(Resources.VerifyVersion_ResultText_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<bool> ResultText { get; set; }

        #endregion


        #region Constructors

        public VerifyVersion()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (ConnectionString == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(ConnectionString)));
            if (BotName == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(BotName)));
            if (BotVersion == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(BotVersion)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var connectionstring = ConnectionString.Get(context);
            var botname = BotName.Get(context);
            var botversion = BotVersion.Get(context);
            string logsessionID;
            string logsysname;
            string logactiondone;
            string logresult;
            string logremarks;
            string logexceptions;
            DateTime dateToday = DateTime.Now;
            string user = System.Environment.UserName;
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(connectionstring);
            string userName =  builder.UserID;
            string validateresult;

            ///////////////////////////
            // Add execution logic HERE

            using (SqlConnection conn = new SqlConnection(connectionstring))
            {
                try
                {
                    //First try-catch wants to validate if database successfully connected.
                    conn.Open();

                    #region declareVariablesForLogs
                    //Insert to System logs - Declare Variables.
                    logsessionID = userName + dateToday.ToString().Replace("/", "");
                    logsessionID = logsessionID.Replace(" ", "");
                    logsessionID = logsessionID.Replace(":", "");
                    var SessionIDtemp = logsessionID;
                    logsysname = userName;
                    logactiondone = "GETTAT - Validate if current solution is updated.";
                    #endregion
                    try
                    {

                        //Execute stored procedure - Validate if bot name is found and bot version is updated
                        SqlCommand cmdvalidateIfUpdatedVersionOut = new SqlCommand("dbo.ValidateIfUpdatedVersion", conn);
                        cmdvalidateIfUpdatedVersionOut.CommandType = CommandType.StoredProcedure;
                        cmdvalidateIfUpdatedVersionOut.Parameters.AddWithValue("@botname", botname);
                        cmdvalidateIfUpdatedVersionOut.Parameters.AddWithValue("@botversion", botversion);
                        SqlDataReader rdrvalidateIfUpdatedVersionOut = cmdvalidateIfUpdatedVersionOut.ExecuteReader();
                        rdrvalidateIfUpdatedVersionOut.Read();
                        validateresult = rdrvalidateIfUpdatedVersionOut.GetValue(0).ToString();


                        if (validateresult == "1")
                        {
                            ResultText.Set(context, true);
                            rdrvalidateIfUpdatedVersionOut.Close();

                            #region InserttoLogs
                            //Declare varables for system logs.
                            logresult = "SUCCESSFUL";
                            logremarks = botname + " - bot is up to date.";
                            logexceptions = "";
                            Console.WriteLine(logremarks);

                            //Insert to System logs.
                            SqlCommand cmdinserttoSystemLogsOut = new SqlCommand("dbo.InsertToSystemLogs", conn);
                            cmdinserttoSystemLogsOut.CommandType = CommandType.StoredProcedure;
                            cmdinserttoSystemLogsOut.Parameters.AddWithValue("@sessionID", SqlDbType.VarChar).Value = SessionIDtemp;
                            cmdinserttoSystemLogsOut.Parameters.AddWithValue("@sysName", SqlDbType.Char).Value = logsysname;
                            cmdinserttoSystemLogsOut.Parameters.AddWithValue("@action", SqlDbType.VarChar).Value = logactiondone;
                            cmdinserttoSystemLogsOut.Parameters.AddWithValue("@result", SqlDbType.VarChar).Value = logresult;
                            cmdinserttoSystemLogsOut.Parameters.AddWithValue("@remarks", SqlDbType.VarChar).Value = logremarks;
                            cmdinserttoSystemLogsOut.Parameters.AddWithValue("@exceptions", SqlDbType.VarChar).Value = logexceptions;
                            cmdinserttoSystemLogsOut.Parameters.AddWithValue("@date", SqlDbType.DateTime).Value = dateToday;
                            cmdinserttoSystemLogsOut.Parameters.AddWithValue("@user", SqlDbType.VarChar).Value = user;
                            cmdinserttoSystemLogsOut.ExecuteNonQuery();
                            #endregion
                        }
                        else
                        {
                            ResultText.Set(context, false);
                            rdrvalidateIfUpdatedVersionOut.Close();

                            #region InserttoLogs
                            //Declare varables for system logs.
                            logresult = "UNSUCCESSFUL";
                            logremarks = botname;
                            logexceptions = "Current version is not updated.";
                            Console.WriteLine(logexceptions);

                            //Insert to System logs.
                            SqlCommand cmdinserttoSystemLogsOut = new SqlCommand("dbo.InsertToSystemLogs", conn);
                            cmdinserttoSystemLogsOut.CommandType = CommandType.StoredProcedure;
                            cmdinserttoSystemLogsOut.Parameters.AddWithValue("@sessionID", SqlDbType.VarChar).Value = SessionIDtemp;
                            cmdinserttoSystemLogsOut.Parameters.AddWithValue("@sysName", SqlDbType.Char).Value = logsysname;
                            cmdinserttoSystemLogsOut.Parameters.AddWithValue("@action", SqlDbType.VarChar).Value = logactiondone;
                            cmdinserttoSystemLogsOut.Parameters.AddWithValue("@result", SqlDbType.VarChar).Value = logresult;
                            cmdinserttoSystemLogsOut.Parameters.AddWithValue("@remarks", SqlDbType.VarChar).Value = logremarks;
                            cmdinserttoSystemLogsOut.Parameters.AddWithValue("@exceptions", SqlDbType.VarChar).Value = logexceptions;
                            cmdinserttoSystemLogsOut.Parameters.AddWithValue("@date", SqlDbType.DateTime).Value = dateToday;
                            cmdinserttoSystemLogsOut.Parameters.AddWithValue("@user", SqlDbType.VarChar).Value = user;
                            cmdinserttoSystemLogsOut.ExecuteNonQuery();
                            #endregion
                        }


                    }
                    catch (Exception e)
                    {
                        ResultText.Set(context, false);

                        #region InserttoLogs
                        //Declare varables for system logs.
                        logresult = "UNSUCCESSFUL";
                        logremarks = botname;
                        logexceptions = "Stored Procedure Error - " + e.ToString();
                        Console.WriteLine(logexceptions);

                        //Insert to System logs.
                        SqlCommand cmdinserttoSystemLogsOut = new SqlCommand("dbo.InsertToSystemLogs", conn);
                        cmdinserttoSystemLogsOut.CommandType = CommandType.StoredProcedure;
                        cmdinserttoSystemLogsOut.Parameters.AddWithValue("@sessionID", SqlDbType.VarChar).Value = SessionIDtemp;
                        cmdinserttoSystemLogsOut.Parameters.AddWithValue("@sysName", SqlDbType.Char).Value = logsysname;
                        cmdinserttoSystemLogsOut.Parameters.AddWithValue("@action", SqlDbType.VarChar).Value = logactiondone;
                        cmdinserttoSystemLogsOut.Parameters.AddWithValue("@result", SqlDbType.VarChar).Value = logresult;
                        cmdinserttoSystemLogsOut.Parameters.AddWithValue("@remarks", SqlDbType.VarChar).Value = logremarks;
                        cmdinserttoSystemLogsOut.Parameters.AddWithValue("@exceptions", SqlDbType.VarChar).Value = logexceptions;
                        cmdinserttoSystemLogsOut.Parameters.AddWithValue("@date", SqlDbType.DateTime).Value = dateToday;
                        cmdinserttoSystemLogsOut.Parameters.AddWithValue("@user", SqlDbType.VarChar).Value = user;
                        cmdinserttoSystemLogsOut.ExecuteNonQuery();
                        #endregion

                    }
                    conn.Close();

                }
                catch (Exception)
                {
                    //Since it's a connection error, Uipath's created activity will throw "false" but will not log the results in the system logs
                    ResultText.Set(context, false);
                }


            }

            ///////////////////////////

            // Outputs
            return (ctx) => {
                //ResultText.Set(ctx, null);
            };
        }

        #endregion
    }
}

