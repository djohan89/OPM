using System;
using System.Data;
using System.Data.SqlClient;

namespace OPM.DBHandler
{

    class OPMDBHandler
    {
        public static string MysqlConnectionString = ConfigurationManager.ConnectionStrings["MyKey"].ConnectionString;
        public static int GetConnection(string query)
        {
            //try 
            //{
                //string strconnection = @"Data Source=DESKTOP-9ILF79D\SQLEXPRESS; Initial Catalog = OPM; User ID = sa; Password=123456";
                //con = new SqlConnection(MysqlConnectionString);
                //return 1;
            //}
            //catch(Exception)
            //{
                //return 0;
            //}
            try
            {
                SqlConnection Conn = new SqlConnection(MysqlConnectionString);
                Conn.Open();
                SqlCommand comm = new SqlCommand(query, Conn);
                comm.ExecuteNonQuery();
                return 1;
                //Conn.Close();
                //MessageBox.Show("Succeeded");
            }
            catch (Exception)
            {
                return 0;
                //MessageBox.Show(ex.Message);
            }
        }

        public static int fInsertData(string strSqlCommand)
        {
            //String strconnection = @"Data Source=DESKTOP-9ILF79D\SQLEXPRESS; Initial Catalog = OPM; User ID = sa; Password=123456";
            SqlConnection con = new SqlConnection(MysqlConnectionString);

            try
            {
                con.Open();
                using (SqlCommand insertCommand = con.CreateCommand())
                {
                    insertCommand.CommandText = strSqlCommand;
                    var row = insertCommand.ExecuteNonQuery();
                }
                CloseConnection(con);
                return 1;
            }
            catch(Exception)
            {
                CloseConnection(con);
                return 0;
            }

        }
        public static int fQuerryData1(string strQuerry)
        {
            //String strconnection = @"Data Source=DESKTOP-9ILF79D\SQLEXPRESS; Initial Catalog = OPM; User ID = sa; Password=123456";
            SqlConnection con = new SqlConnection(MysqlConnectionString);
            try
            {
                con.Open();
                using (SqlCommand sCommand = con.CreateCommand())
                {
                    sCommand.CommandText = strQuerry;
                    using (SqlDataReader reader = sCommand.ExecuteReader())
                    {
                        if(reader.HasRows)
                        {
                            while(reader.Read())
                            {
                                //Do Sothing
                            }    
                        }
                        else
                        {
                            //Do Something
                        }    

                    }    
                }
                CloseConnection(con);
                return 1;

            }
            catch (Exception)
            {
                CloseConnection(con);
                return 0;
            }
        }

        public static int fQuerryData(string strQuerry, ref DataSet ds)
        {
            //String strconnection = @"Data Source=DESKTOP-9ILF79D\SQLEXPRESS; Initial Catalog = OPM; User ID = sa; Password=123456";
            SqlConnection con = new SqlConnection(MysqlConnectionString);
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommand command;

            try
            {
                con.Open();
                command = new SqlCommand(strQuerry, con);
                adapter.SelectCommand = command;
                adapter.Fill(ds, "SQL Temp");

                adapter.Dispose();
                command.Dispose();
                CloseConnection(con);
                if(0 == ds.Tables[0].Rows.Count)
                {
                    return 0;
                }    
                return 1;

            }
            catch (Exception)
            {
                adapter.Dispose();
                CloseConnection(con);
                return 0;
            }
        }
        public static int CloseConnection(SqlConnection con)
        {
            con.Close();
            con.Dispose();
            con = null;
            return 1;
        }
    }
}
