﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using Oracle.ManagedDataAccess.Client;
using System.Data;
//using System.Data.OracleClient;

namespace ExchangeRecon
{
    class DBConnector : IDisposable
    {
        private static readonly string Hostname = ConfigurationManager.AppSettings["DBHostName"];
        private static readonly string Port = ConfigurationManager.AppSettings["DBPort"];
        private static readonly string Username = ConfigurationManager.AppSettings["DBUserName"];
        private static readonly string Pwd = ConfigurationManager.AppSettings["DBPassword"];
        private static readonly string SID = ConfigurationManager.AppSettings["DBSID"];
        private readonly OracleConnection _connection = null;

        ~DBConnector()
        {
            Dispose();
        }

        public DBConnector()
        {
            _connection = BuildConnection();
        }
        private static OracleConnection BuildConnection()
        {
            string connectionString = string.Format(
                                       ConfigurationManager.ConnectionStrings["OracleConnectionString"].ConnectionString,
                                       Username,
                                       Pwd,
                                       Hostname,
                                       Port.ToString(),
                                       SID
                );
            OracleConnection connection = new OracleConnection { ConnectionString = connectionString };

            try
            {
                connection.Open();
            }
            catch (OracleException ex)
            {
                //Logger.Error(ex.ToString());
                connection.Dispose();
                connection = null;
                Console.WriteLine(ex.ToString());
                //throw ex;
            }
            return connection;
        }
        public DataTable MakeQuery(string query, string tableName)
        {
            OracleCommand cmd = new OracleCommand();
            DataTable dataTable = new DataTable();
            dataTable.TableName = tableName;
            try
            {
                if (_connection != null)
                {
                    cmd.Connection = _connection;
                    cmd.CommandText = query;
                    cmd.CommandType = CommandType.Text;
                    OracleDataReader reader = cmd.ExecuteReader();
                    var columns = new List<string>();
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        string col = reader.GetName(i);
                        columns.Add(col);
                        dataTable.Columns.Add(col);
                    }
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            var row = dataTable.NewRow();
                            foreach (string col in columns)
                            {
                                row[col] = reader[col];
                            }
                            dataTable.Rows.Add(row);
                        }
                    }
                    return dataTable;
                }

                else
                {
                    throw new MissingMemberException("Database Connection failed");
                }
            }
            catch (Exception ex)
            {
                throw new OracleQueryAbortException(ex.ToString());
            }

        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                    _connection?.Dispose();
                    _connection?.Close();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~DBConnector() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion

    }
}
