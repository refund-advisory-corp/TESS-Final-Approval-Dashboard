using System;
using System.Collections.Generic;
using System.Linq;
using System.Collections.ObjectModel;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Threading;

namespace TESS_Dashboard
{
    public class SQLcommander
    {
        private SqlConnection ExistingConnection;

        public SQLcommander(SqlConnection ExistingConnectionInput = null)
        {
            ConnectionString = ConfigurationManager.ConnectionStrings["TESS Dashboard.Properties.Settings.REFTBL4ConnectionString"].ConnectionString; //
            ExistingConnection = ExistingConnectionInput;
        }

        public void Select()
        {

            string SqlText = "SELECT ";

            int counter = 0;
            foreach (string s in Requests)
            {
                if (counter > 0)
                {
                    SqlText += ", " + s;
                }
                else
                {
                    SqlText += s;
                }
                counter += 1;
            }

            SqlText += " FROM " + TableName + " WHERE ";

            counter = 0;
            foreach (string s in Conditions)
            {
                if (counter > 0)
                {
                    SqlText += " AND " + s;
                }
                else
                {
                    SqlText += s;
                }
                counter += 1;
            }

            int RecordCount = Requests[0].IndexOf("TOP 1") > -1 ? 1 : CountRecords();


            SqlConnection connection;
            if (ExistingConnection != null)
            {
                connection = ExistingConnection;
            }
            else
            {
                connection = new SqlConnection(ConnectionString);
                connection.Open();
            }

            object[,] SelectReturn = new object[RecordCount, Requests.Count];

            //MessageBox.Show("SELECT command running: " + Environment.NewLine + SqlText);

            using (SqlCommand cmd = new SqlCommand(SqlText, connection))
            {
                counter = 0;
                foreach (string s in Parameters.Keys)
                {
                    cmd.Parameters.AddWithValue(s, Parameters[s]);
                    counter += 1;
                }

                SqlDataReader reader = cmd.ExecuteReader();

                counter = 0;
                while (reader.Read())
                {

                    for (int i = 0; i < Requests.Count; i++)
                    {
                        SelectReturn[counter, i] = (reader.GetValue(i).GetType() == typeof(DBNull)) ? null : reader.GetValue(i);
                    }
                    counter += 1;
                }
                reader.Close();
            }

            if (ExistingConnection == null)
            {
                connection.Close();
            }
            

            SelectResults = SelectReturn;
        }

        public void Update()
        {
            string SqlText = "UPDATE " + TableName + " SET ";

            int counter = 0;
            foreach (string s in Requests)
            {
                if (counter > 0)
                {
                    SqlText += ", " + s;
                }
                else
                {
                    SqlText += s;
                }
                counter += 1;
            }

            SqlText += " WHERE ";
            counter = 0;
            foreach (string s in Conditions)
            {
                if (counter > 0)
                {
                    SqlText += " AND " + s;
                }
                else
                {
                    SqlText += s;
                }
                counter += 1;
            }

            //MessageBox.Show("UPDATE command running: " + Environment.NewLine + SqlText);
            SqlConnection connection;
            if (ExistingConnection != null)
            {
                connection = ExistingConnection;
            }
            else
            {
                connection = new SqlConnection(ConnectionString);
                connection.Open();
            }

            using (SqlCommand cmd = new SqlCommand(SqlText, connection))
            {
                counter = 0;
                foreach (string s in Parameters.Keys)
                {
                    cmd.Parameters.AddWithValue(s, Parameters[s]);
                    counter += 1;
                }

                RowsAffected = cmd.ExecuteNonQuery();
            }

            if (ExistingConnection == null)
            {
                connection.Close();
            }
            

        }

        public void Insert()
        {
            string SqlText = "INSERT INTO " + TableName + " (";

            int counter = 0;
            foreach (string s in Requests)
            {
                if (counter > 0)
                {
                    SqlText += ", " + s;
                }
                else
                {
                    SqlText += s;
                }
                counter += 1;
            }

            SqlText += ") VALUES (";

            counter = 0;
            foreach (string s in Parameters.Keys)
            {
                if (counter > 0)
                {
                    SqlText += ", " + s;
                }
                else
                {
                    SqlText += s;
                }
                counter += 1;
            }

            SqlText += ")";

            SqlConnection connection;
            if (ExistingConnection != null)
            {
                connection = ExistingConnection;
            }
            else
            {
                connection = new SqlConnection(ConnectionString);
                connection.Open();
            }

            //MessageBox.Show("INSERT command running: " + Environment.NewLine + SqlText);

            using (SqlCommand cmd = new SqlCommand(SqlText, connection))
            {
                counter = 0;
                foreach (string s in Parameters.Keys)
                {
                    cmd.Parameters.AddWithValue(s, Parameters[s]);
                    counter += 1;
                }

                RowsAffected = cmd.ExecuteNonQuery();
            }

            if (ExistingConnection == null)
            {
                connection.Close();
            }
        }

        public int CountRecords()
        {
            string SqlCount = "SELECT COUNT(*)";

            SqlCount += " FROM " + TableName + " WHERE ";

            int counter = 0;
            foreach (string s in Conditions)
            {
                if (counter > 0)
                {
                    SqlCount += " AND " + s;
                }
                else
                {
                    SqlCount += s;
                }
                counter += 1;
            }

            int CountRecordsReturn;
            SqlConnection connection;
            if (ExistingConnection != null)
            {
                connection = ExistingConnection;
            }
            else
            {
                connection = new SqlConnection(ConnectionString);
                connection.Open();
            }

            using (SqlCommand cmd = new SqlCommand(SqlCount, connection))
            {

                counter = 0;
                foreach (string s in Parameters.Keys)
                {
                    cmd.Parameters.AddWithValue(s, Parameters[s]);
                    counter += 1;
                }
                CountRecordsReturn = (int)cmd.ExecuteScalar();
            }

            if (ExistingConnection == null)
            {
                connection.Close();
            }
            return CountRecordsReturn;

        }

        public void AddCondition(string Condition)
        {
            Conditions.Add(Condition);
        } //MPA 8/15/2019 Deprecated

        public void AddParameter(string title, object parameter) //MPA 8/15/2019 Deprecated
        {
            //ParameterTitles.Add(title);
            //Parameters.Add(parameter);
        }

        public void AddRequest(string Request)
        {
            Requests.Add(Request);
        } //MPA 8/15/2019 Deprecated

        public dynamic GetRequestData(string Request)
        {
            for (int i = 0; i < Requests.Count; i++)
            {

                if (Request == Requests[i])
                {
                    return RowFromMDA(SelectResults, i); //selects the corresponding array
                }

            }
            object[] emptyreturn = new object[0];
            return emptyreturn; //only reached if other return does not occur

        }

        public object[] RowFromMDA(object[,] MDA, int rowindex)
        {
            int columns = MDA.GetLength(0);
            int rows = MDA.GetLength(1);

            object[] RowFromMDAReturn = new object[columns];
            for (int i = 0; i < columns; i++)
            {
                RowFromMDAReturn[i] = MDA[i, rowindex];

            }

            return RowFromMDAReturn;

        }

        public Collection<string> Conditions { get; set; } = new Collection<string>();

        //public Collection<string> ParameterTitles { get; set; } = new Collection<string>();
        //public Collection<object> Parameters { get; set; } = new Collection<object>();

        public Dictionary<string, object> Parameters { get; set; } = new Dictionary<string, object>();
        public Collection<string> Requests { get; set; } = new Collection<string>();
        public object[,] SelectResults { get; set; }

        public int RowsAffected { get; set; } = 0;

        public string ConnectionString { get; set; }
        public string TableName { get; set; }
    }
}
