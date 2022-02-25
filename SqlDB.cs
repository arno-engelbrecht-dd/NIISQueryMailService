using System;
using System.Data;
using System.Data.SqlClient;

namespace NIISQueryMailService
{
    public class SqlDB : IDisposable
    {
        private SqlConnection DBConnection;
        private static String m_DBConnectionString = "";
        private static readonly object m_Lock = new object();
        private bool _disposed;
        private bool _intransaction;

        public static string strConnection
        {
            get
            {
                if (string.IsNullOrEmpty(m_DBConnectionString))
                {
                    var appSettings = new System.Configuration.AppSettingsReader();
                    m_DBConnectionString = (string)(appSettings.GetValue("DBConnectionString", typeof(string)));
                }
                return m_DBConnectionString;
            }
        }

        public SqlDB()
        {
            DBConnection = new SqlConnection(strConnection);
            _disposed = false;
            _intransaction = false;
        }

        public SqlDB(string pConnection)
        {
            DBConnection = new SqlConnection(pConnection);
            _disposed = false;
            _intransaction = false;
        }

        public void Dispose()
        {
            try
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
            catch (Exception)
            {
            }
        }

        private void OpenConnection()
        {
            if (DBConnection.State != ConnectionState.Open)
            {
                lock (m_Lock)
                {
                    DBConnection.Open();
                }
            }
        }

        public Object ReadSingleValue(String query, Object[] cmdParams)
        {
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand(query, DBConnection);
                OpenConnection();
                if (cmdParams != null)
                {
                    for (int i = 1, j = cmdParams.Length; i <= j; i++)
                    {
                        cmd.Parameters.AddWithValue("@" + i, cmdParams[i - 1]);
                    }
                }
                Object obj = cmd.ExecuteScalar();
                //DBConnection.Close();
                if (obj == null)
                    return "";
                return obj;
            }
            finally
            {
                if (cmd != null)
                    cmd.Dispose();
            }
        }

        public int ReadInt(String query, Object[] cmdParams)
        {
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand(query, DBConnection);
                OpenConnection();
                if (cmdParams != null)
                {
                    for (int i = 1, j = cmdParams.Length; i <= j; i++)
                    {
                        cmd.Parameters.AddWithValue("@" + i, cmdParams[i - 1]);
                    }
                }
                Object obj = cmd.ExecuteScalar();
                if (obj == null)
                    return 0;
                int iVal = 0;
                if (Int32.TryParse(obj.ToString(), out iVal))
                    return iVal;

                return 0;
            }
            finally
            {
                if (cmd != null)
                    cmd.Dispose();
            }
        }

        public SqlDataReader RunQuery(String query, Object[] cmdParams)
        {
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand(query, DBConnection);
                OpenConnection();
                if (cmdParams != null)
                {
                    for (int i = 1, j = cmdParams.Length; i <= j; i++)
                    {
                        cmd.Parameters.AddWithValue("@" + i, cmdParams[i - 1]);
                    }
                }
                SqlDataReader reader = cmd.ExecuteReader();
                return reader;
            }
            finally
            {
                if (cmd != null)
                    cmd.Dispose();
            }
        }

        public SqlDataReader RunLongQuery(String query, Object[] cmdParams)
        {
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand(query, DBConnection);
                cmd.CommandTimeout = 600;
                OpenConnection();
                if (cmdParams != null)
                {
                    for (int i = 1, j = cmdParams.Length; i <= j; i++)
                    {
                        cmd.Parameters.AddWithValue("@" + i, cmdParams[i - 1]);
                    }
                }
                SqlDataReader reader = cmd.ExecuteReader();
                return reader;
            }
            finally
            {
                if (cmd != null)
                    cmd.Dispose();
            }
        }

        public void ExecuteSQL(String query, Object[] cmdParams)
        {
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand(query, DBConnection);
                OpenConnection();
                if (cmdParams != null)
                {
                    for (int i = 1, j = cmdParams.Length; i <= j; i++)
                    {
                        cmd.Parameters.AddWithValue("@" + i, cmdParams[i - 1]);
                    }
                }
                Object obj = cmd.ExecuteNonQuery();
            }
            finally
            {
                if (cmd != null)
                    cmd.Dispose();
                DBConnection.Close();
            }
        }

        public void ExecuteSQL_NoClose(String query, Object[] cmdParams)
        {
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand(query, DBConnection);
                OpenConnection();
                if (cmdParams != null)
                {
                    for (int i = 1, j = cmdParams.Length; i <= j; i++)
                    {
                        cmd.Parameters.AddWithValue("@" + i, cmdParams[i - 1]);
                    }
                }
                Object obj = cmd.ExecuteNonQuery();
            }
            finally
            {
                if (cmd != null)
                    cmd.Dispose();
            }
        }

        public void ExecuteLongSQL(String query, Object[] cmdParams)
        {
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand(query, DBConnection);
                cmd.CommandTimeout = 600;
                OpenConnection();
                if (cmdParams != null)
                {
                    for (int i = 1, j = cmdParams.Length; i <= j; i++)
                    {
                        cmd.Parameters.AddWithValue("@" + i, cmdParams[i - 1]);
                    }
                }
                Object obj = cmd.ExecuteNonQuery();
            }
            finally
            {
                if (cmd != null)
                    cmd.Dispose();
            }
        }

        public DataSet FillDataSet(String query, String Name, Object[] cmdParams)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = null;
            SqlDataAdapter adapter = null;
            try
            {
                cmd = new SqlCommand(query, DBConnection);
                if (cmdParams != null)
                {
                    for (int i = 1, j = cmdParams.Length; i <= j; i++)
                    {
                        cmd.Parameters.AddWithValue("@" + i, cmdParams[i - 1]);
                    }
                }
                OpenConnection();
                adapter = new SqlDataAdapter(cmd);
                adapter.Fill(ds, Name);
            }
            finally
            {
                if (cmd != null)
                    cmd.Dispose();
                if (adapter != null)
                    adapter.Dispose();
                DBConnection.Close();
            }
            return ds;
        }

        public SqlCommand CreateCommand(String query)
        {
            OpenConnection();

            return (new SqlCommand(query, DBConnection));
        }

        public void ExecuteCommand(SqlCommand cmd)
        {
            OpenConnection();
            cmd.ExecuteNonQuery();
        }


        public Object ExecuteQueryCommand(SqlCommand cmd)
        {
            OpenConnection();
            Object obj = cmd.ExecuteScalar();
            if (obj == null)
                return string.Empty;
            return obj;
        }

        public void BeginTransaction()
        {
            try
            {
                ExecuteSQL("BEGIN TRANSACTION", null);
                _intransaction = true;
            }
            catch { }
        }

        public void Commit()
        {
            if (_intransaction)
            {
                _intransaction = false;
                try
                {
                    ExecuteSQL("COMMIT", null);
                }
                catch { }
            }
        }

        public void Rollback()
        {
            if (_intransaction)
            {
                _intransaction = false;
                try
                {
                    ExecuteSQL("ROLLBACK", null);
                }
                catch { }
            }
        }

        public void CloseConnection()
        {
            try
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
            catch (Exception)
            {
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            lock (m_Lock)
            {
                if (_disposed == false)
                {
                    if (disposing)
                    {
                        if (DBConnection != null)
                        {
                            // Commit any changes
                            if (_intransaction)
                            {
                                if (DBConnection.State == ConnectionState.Open)
                                {
                                    Commit();
                                }
                            }

                            if (DBConnection.State != ConnectionState.Closed)
                            {
                                DBConnection.Close();
                            }

                            DBConnection.Dispose();
                            DBConnection = null;
                        }
                        _disposed = true;
                    }
                }
            }
        }
    }
}