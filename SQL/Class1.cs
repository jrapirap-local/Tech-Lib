
using Microsoft.Data.SqlClient;
using System.Data;

namespace SQL
{
    public static class sql
    {
        private static readonly SqlConnection _con = new SqlConnection();
        public static void Open(string con)
        { //
            _con.ConnectionString = con;
            _con.Open();
        }

        public static void ExecuteNonQuery(string cmd)
        {
            new SqlCommand(cmd, _con).ExecuteNonQuery();
        }

        //Overloading
        public static void ExecuteNonQuery(string cmd, params SqlParameter[] param)
        {
            using (var com = new SqlCommand(cmd, _con))
            {
                com.CommandType = CommandType.StoredProcedure;
                com.Parameters.AddRange(param);
                com.ExecuteNonQuery();
            }
        }

        public static DataTable GetDataTable(string cmd)
        {
            using (DataTable tbl = new DataTable())
            {
                new SqlDataAdapter(cmd, _con).Fill(tbl);
                return tbl;
            }
        }

        public static DataSet GetDataSet(string cmd)
        {
            using (DataSet tbl = new DataSet())
            {
                new SqlDataAdapter(cmd, _con).Fill(tbl);
                return tbl;
            }
        }

        public static DataTable GetDataTable(string cmd, params SqlParameter[] param)
        {
            using (DataTable tbl = new DataTable())
            {
                using (var com = new SqlCommand(cmd, _con))
                {
                    com.CommandType = CommandType.StoredProcedure;
                    com.Parameters.AddRange(param);
                    new SqlDataAdapter(com).Fill(tbl);
                    return tbl;
                }
            }
        }

        public static DataSet GetDataSet(string cmd, params SqlParameter[] param)
        {
            using (DataSet tbl = new DataSet())
            {
                using (var com = new SqlCommand(cmd, _con))
                {
                    com.CommandType = CommandType.StoredProcedure;
                    com.Parameters.AddRange(param);
                    new SqlDataAdapter(com).Fill(tbl);
                    return tbl;
                }
            }
        }


        public static dynamic ExecuteScalar(string cmd)
        {
            return new SqlCommand(cmd, _con).ExecuteScalar();
        }

        public static dynamic ExecuteScalar(string cmd, params SqlParameter[] param)
        {
            using (var com = new SqlCommand(cmd, _con))
            {
                com.CommandType = CommandType.StoredProcedure;
                com.Parameters.AddRange(param);
                return com.ExecuteScalar();
            }
        }
    }
}
