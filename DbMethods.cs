//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

// OleDB Database methods

namespace SharpPlant
{
    public static class DbMethods
    {
        public static DataTable GetDbTable(string dbPath, string tableName)
        {
            using (var connection = GetConnection(dbPath, ConnectionType.Ace))
            {
                CheckOpenConnection(connection);

                var returnTable = new DataTable(tableName);
                var selectCommand = GetSelectCommand(tableName, connection);

                selectCommand.CommandType = CommandType.TableDirect;

                var adapter = new OleDbDataAdapter(selectCommand);
                adapter.FillSchema(returnTable, SchemaType.Mapped);
                adapter.Fill(returnTable);

                return returnTable;
            }
        }

        public static Image GetDbImage(object imageField)
        {
            if (imageField == DBNull.Value)
                return null;

            // Image header constants
            const string BITMAP_ID_BLOCK = "BM";
            const string JPG_ID_BLOCK = "\u00FF\u00D8\u00FF";
            const string PNG_ID_BLOCK = "\u0089PNG\r\n\u001a\n";
            const string GIF_ID_BLOCK = "GIF8";
            const string TIFF_ID_BLOCK = "II*\u0000";

            var buffer = (byte[])imageField;
            var e7 = Encoding.UTF7;
            var byteString = e7.GetString(buffer).Substring(0, 300);
            int iPos = -1;

            // Bitmap
            if (byteString.IndexOf(BITMAP_ID_BLOCK) != -1)
                iPos = byteString.IndexOf(BITMAP_ID_BLOCK);

            // Jpeg
            else if (byteString.IndexOf(JPG_ID_BLOCK) != -1)
                iPos = byteString.IndexOf(JPG_ID_BLOCK);

            // Png
            else if (byteString.IndexOf(PNG_ID_BLOCK) != -1)
                iPos = byteString.IndexOf(PNG_ID_BLOCK);

            // Gif
            else if (byteString.IndexOf(GIF_ID_BLOCK) != -1)
                iPos = byteString.IndexOf(GIF_ID_BLOCK);

            // Tiff
            else if (byteString.IndexOf(TIFF_ID_BLOCK) != -1)
                iPos = byteString.IndexOf(TIFF_ID_BLOCK);

            else
                throw new Exception("Unable to determine header size for the OLE Object");

            using (var imgStream = new MemoryStream())
            {
                imgStream.Write(buffer, iPos, buffer.Length - iPos);
                return Image.FromStream(imgStream);
            }
        }

        public static bool UpdateDbTable(string dbPath, DataTable inputTable)
        {
            using (var connection = GetConnection(dbPath))
            {
                try
                {
                    CheckOpenConnection(connection);

                    var commandString = string.Format("SELECT * FROM {0}", inputTable.TableName);
                    var dataAdapter = new OleDbDataAdapter(commandString, connection);

                    foreach (var row in from DataRow row in inputTable.Rows
                                        where row.RowState != DataRowState.Unchanged
                                        let command = new OleDbCommandBuilder(dataAdapter).GetUpdateCommand(true)
                                        select row)
                    {
                        dataAdapter.Update(new[] { row });
                    }

                    inputTable.AcceptChanges();
                    return true;

                }
                catch (OleDbException ex)
                {
                    Debug.WriteLine(ex.Message);
                    throw;
                }
            }
        }
        public static bool UpdateDbTable(string dbPath, string rowFilter, DataTable inputTable)
        {
            using (var connection = GetConnection(dbPath))
            {
                try
                {
                    CheckOpenConnection(connection);

                    var updateCommand = GetUpdateCommand(inputTable, rowFilter, connection);
                    var adapter = new OleDbDataAdapter { UpdateCommand = updateCommand };

                    adapter.Update(inputTable);

                    inputTable.AcceptChanges();

                }
                catch (OleDbException ex)
                {
                    Debug.WriteLine(ex.Message);
                    throw;
                }
            }

            // Return true on success
            return true;
        }

        public static bool AddDbField(string dbPath, string tableName, string fieldName)
        {
            return AddDbField(dbPath, tableName, fieldName, "TEXT(255)");
        }
        public static bool AddDbField(string dbPath, string tableName, string fieldName, string fieldType)
        {
            using (var connection = GetConnection(dbPath))
            {
                try
                {
                    CheckOpenConnection(connection);
                    var updateCommand = GetAddFieldCommand(tableName, fieldName, fieldType, connection);
                    updateCommand.ExecuteNonQuery();

                    return true;
                }

                catch (OleDbException ex)
                {
                    if (ex.ErrorCode == -2147217887)
                        return true; // field exists

                    throw;
                }
            }
        }

        internal static void CheckOpenConnection(OleDbConnection connection)
        {
            try { connection.Open(); }
            catch (InvalidOperationException)
            {
                try
                {
                    switch (connection.Provider)
                    {
                        case "Microsoft.Ace.OLEDB.12.0":
                            connection = GetConnection(connection.DataSource, ConnectionType.Jet);
                            break;
                        case "Microsoft.Jet.OLEDB.4.0":
                            connection = GetConnection(connection.DataSource, ConnectionType.Ace);
                            break;
                    }

                    connection.Open();
                }
                catch (InvalidOperationException)
                {
                    throw new Exception("Connection could not be established.  OleDB Ace/Jet drivers not installed.");
                }
            }
        }

        internal static OleDbConnection GetConnection(string dbPath)
        {
            return GetConnection(dbPath, ConnectionType.Ace);
        }
        internal static OleDbConnection GetConnection(string dbPath, ConnectionType type)
        {
            string connectionString;
            switch (type)
            {
                case ConnectionType.Ace:
                    connectionString = string.Format("Provider=Microsoft.Ace.OLEDB.12.0;Data Source ={0};", dbPath);
                    break;
                case ConnectionType.Jet:
                    connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source ={0};", dbPath);
                    break;
                default:
                    return null;
            }

            return new OleDbConnection(connectionString);
        }

        private static OleDbCommand GetSelectCommand(string tableName, OleDbConnection connection)
        {
            var retCommand = connection.CreateCommand();
            retCommand.CommandText = string.Format("SELECT * FROM {0}", tableName);

            return retCommand;
        }

        private static OleDbCommand GetUpdateCommand(DataTable inputTable, string rowFilter, OleDbConnection connection)
        {
            var retCommand = connection.CreateCommand();
            var sb = new StringBuilder(string.Format("UPDATE {0} SET ", inputTable.TableName));

            foreach (DataColumn col in inputTable.Columns)
            {
                sb.AppendFormat("{0} = ?, ", col.ColumnName);
                var par = new OleDbParameter
                {
                    ParameterName = "@" + col.ColumnName,
                    OleDbType = GetOleDbType(col.DataType),
                    Size = col.MaxLength,
                    SourceColumn = col.ColumnName,
                };

                retCommand.Parameters.Add(par);
            }

            sb.Remove(sb.ToString().LastIndexOf(','), 1);

            // Add a where clause if a rowfilter was provided
            if (rowFilter != string.Empty)
                sb.AppendFormat("WHERE {0}", rowFilter);

            retCommand.CommandText = sb.ToString();
            return retCommand;
        }
        private static OleDbCommand GetUpdateCommand(DataRow row, OleDbConnection connection)
        {
            var retCommand = connection.CreateCommand();
            var parentTable = row.Table;
            var pKey = parentTable.PrimaryKey[0];

            var sb = new StringBuilder(string.Format("UPDATE {0} SET ", parentTable.TableName));

            foreach (var col in parentTable.Columns
                    .Cast<DataColumn>()
                    .Where(col => !col.Unique))
            {
                sb.AppendFormat("{0} = ?, ", col.ColumnName);

                var par = new OleDbParameter
                {
                    ParameterName = col.ColumnName,
                    OleDbType = GetOleDbType(col.DataType),
                    Size = col.MaxLength,
                    SourceColumn = col.ColumnName,
                };

                retCommand.Parameters.Add(par);
            }

            sb.Remove(sb.ToString().LastIndexOf(','), 1);

            // Add a where clause to the primary key
            sb.AppendFormat("WHERE {0} = {1}", pKey.ColumnName, row[pKey.ColumnName]);

            retCommand.CommandText = sb.ToString();
            return retCommand;
        }

        private static OleDbCommand GetAddFieldCommand(string tableName, string fieldName, string fieldType, OleDbConnection connection)
        {
            var retCommand = connection.CreateCommand();
            var sb = new StringBuilder(string.Format("ALTER TABLE {0} ADD COLUMN ", tableName));

            fieldName = fieldName.Replace(' ', '_');

            sb.AppendFormat("{0} {1}, ", fieldName, fieldType);
            sb.Remove(sb.ToString().LastIndexOf(','), 1);

            retCommand.CommandText = sb.ToString();
            return retCommand;
        }

        private static OleDbType GetOleDbType(Type inputType)
        {
            switch (inputType.FullName)
            {
                // Return the appropriate type
                case "System.Boolean":
                    return OleDbType.Boolean;
                case "System.Int32":
                    return OleDbType.Integer;
                case "System.Single":
                    return OleDbType.Single;
                case "System.Double":
                    return OleDbType.Double;
                case "System.Decimal":
                    return OleDbType.Decimal;
                case "System.String":
                    return OleDbType.Char;
                case "System.Char":
                    return OleDbType.Char;
                case "System.Byte[]":
                    return OleDbType.Binary;
                default:
                    return OleDbType.Variant;
            }
        }

        public enum ConnectionType
        {
            Ace = 1,
            Jet = 2
        }
    }
}//