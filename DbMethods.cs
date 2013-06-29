//
//  Copyright © 2013 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Text;

// OleDB Database methods

namespace SharpPlant
{
    public static class DbMethods
    {
        public static DataTable GetDbTable(string dbPath, string tableName)
        {
            // Create the return table
            var returnTable = new DataTable(tableName);

            // Set the initial connection type
            var conType = ConnectionType.Ace;

            ConnectionStart:

            // Create the connection
            using (var connection = GetConnection(dbPath, conType))
            {
                try
                {
                    // Open the connection
                    connection.Open();

                    // Build the select command
                    var selectCommand = GetSelectCommand(tableName, connection);

                    // Create the data adapter
                    var adapter = new OleDbDataAdapter(selectCommand);

                    // Fill the return table
                    adapter.Fill(returnTable);
                }

                // Change the connection type on error
                catch (InvalidOperationException)
                {
                    // Check if JET drivers have been attempted
                    if (conType == ConnectionType.Jet)
                    {
                        throw new Exception("Connection could not be established.  OleDB Ace/Jet drivers not installed.");
                    }

                    // Set the connection type to JET drivers and retry the connection
                    conType = ConnectionType.Jet;
                    goto ConnectionStart;
                }

                // Return nothing on error
                catch (OleDbException ex)
                {
                    throw new Exception(ex.Message);
                }
            }

            // Return the table
            return returnTable;
        }

        public static Image GetDbImage(object imageField)
        {
            // Check for DbNull
            if (imageField == DBNull.Value)
                return null;

            // Image header constants
            const string BITMAP_ID_BLOCK = "BM";
            const string JPG_ID_BLOCK = "\u00FF\u00D8\u00FF";
            const string PNG_ID_BLOCK = "\u0089PNG\r\n\u001a\n";
            const string GIF_ID_BLOCK = "GIF8";
            const string TIFF_ID_BLOCK = "II*\u0000";

            // Cast the field to a byte array
            var buffer = (byte[])imageField;

            // Set the default encoding
            var e7 = Encoding.UTF7;

            // Get the bytes as a string, and set the starting position
            String byteString = e7.GetString(buffer).Substring(0, 300);
            int iPos = -1;

            // Check for a bitmap
            if (byteString.IndexOf(BITMAP_ID_BLOCK) != -1)
                iPos = byteString.IndexOf(BITMAP_ID_BLOCK);

            // Check for a jpeg
            else if (byteString.IndexOf(JPG_ID_BLOCK) != -1)
                iPos = byteString.IndexOf(JPG_ID_BLOCK);

            // Check for a png
            else if (byteString.IndexOf(PNG_ID_BLOCK) != -1)
                iPos = byteString.IndexOf(PNG_ID_BLOCK);

            // Check for a gif
            else if (byteString.IndexOf(GIF_ID_BLOCK) != -1)
                iPos = byteString.IndexOf(GIF_ID_BLOCK);

            // Check for a tiff
            else if (byteString.IndexOf(TIFF_ID_BLOCK) != -1)
                iPos = byteString.IndexOf(TIFF_ID_BLOCK);

            // Throw an exception if the image tye was undetermined
            else throw new Exception("Unable to determine header size for the OLE Object");
            if (iPos == -1) throw new Exception("Unable to determine header size for the OLE Object");

            // Load the image into memory and return it
            using (var imgStream = new MemoryStream())
            {
                imgStream.Write(buffer, iPos, buffer.Length - iPos);
                return Image.FromStream(imgStream);
                //return stream.ToArray();
            }
        }

        public static bool UpdateDbTable(string dbPath, DataTable inputTable)
        {
            return UpdateDbTable(dbPath, string.Empty, inputTable);
        }

        public static bool UpdateDbTable(string dbPath, string rowFilter, DataTable inputTable)
        {
            // Create the connection
            using (var connection = GetConnection(dbPath))
            {
                try
                {
                    // Open the connection
                    connection.Open();

                    // Create the update command
                    var updateCommand = GetUpdateCommand(inputTable, rowFilter, connection);

                    // Link the new data adapter
                    var adapter = new OleDbDataAdapter { UpdateCommand = updateCommand };

                    // Update the MDB table
                    adapter.Update(inputTable);

                } // Return false on error
                catch (OleDbException ex)
                {
                    throw ex;
                    //return false;
                }
            }

            // Return true on success
            return true;
        }

        public static bool DeleteDbValue(string dbpath, string tablename, string fieldname)
        {
            return false;
        }

        public static bool AddDbField(string dbPath, string tableName, string fieldName)
        {
            return AddDbField(dbPath, tableName, fieldName, "TEXT(255)");
        }
        public static bool AddDbField(string dbPath, string tableName, string fieldName, string fieldType)
        {
            // Create the connection
            using (var connection = GetConnection(dbPath))
            {
                try
                {
                    // Open the connection
                    connection.Open();

                    // Create the update command
                    var updateCommand = GetAddFieldCommand(tableName, fieldName, fieldType, connection);

                    // Update the database table 
                    updateCommand.ExecuteNonQuery();
                }

                // Return false on error
                catch (OleDbException ex)
                {
                    // Return true if the field already exists, otherwise false
                    return ex.ErrorCode == -2147217887;
                }
            }

            // Return true on success
            return true;
        }

        private static OleDbConnection GetConnection(string dbPath)
        {
            return GetConnection(dbPath, ConnectionType.Ace);
        }
        private static OleDbConnection GetConnection(string dbPath, ConnectionType type)
        {
            // Build the connection string
            string connectionString;
            if (type == ConnectionType.Ace)
            {
                connectionString = string.Format("Provider=Microsoft.Ace.OLEDB.12.0;Data Source ={0};", dbPath);
            }
            else
            {
                connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source ={0};", dbPath);
            }

            // Return the connection
            return new OleDbConnection(connectionString);
        }

        private static OleDbCommand GetSelectCommand(string tableName, OleDbConnection connection)
        {
            // Create the return command
            var retCommand = connection.CreateCommand();

            // Build the command string
            retCommand.CommandText = string.Format("SELECT * FROM {0}", tableName);

            // Return the command
            return retCommand;
        }

        private static OleDbCommand GetUpdateCommand(DataTable inputTable, string rowFilter, OleDbConnection connection)
        {
            // Create the return command
            var retCommand = connection.CreateCommand();

            // Build the command string
            var sb = new StringBuilder(string.Format("UPDATE {0} SET ", inputTable.TableName));

            foreach (DataColumn col in inputTable.Columns)
            {
                // Append the command text
                sb.AppendFormat("{0} = ?, ", col.ColumnName);

                // Create the column parameter
                var par = new OleDbParameter
                    {
                        ParameterName = col.ColumnName,
                        OleDbType = GetOleDbType(col.DataType),
                        Size = col.MaxLength,
                        SourceColumn = col.ColumnName,
                    };
                    
                // Add the parameter to the return command
                retCommand.Parameters.Add(par);
            }

            // Remove the last comma
            sb.Remove(sb.ToString().LastIndexOf(','), 1);

            // Add a where clause if a rowfilter was provided
            if (rowFilter != string.Empty)
                sb.AppendFormat("WHERE {0}", rowFilter);

            // Set the command text
            retCommand.CommandText = sb.ToString();

            // Return the command
            return retCommand;
        }

        private static OleDbCommand GetAddFieldCommand(string tableName, string fieldName, string fieldType, OleDbConnection connection)
        {
            // Create the return command
            var retCommand = connection.CreateCommand();

            // Build the command string
            var sb = new StringBuilder(string.Format("ALTER TABLE {0} ADD COLUMN ", tableName));

            // Replace any spaces from the field name
            fieldName = fieldName.Replace(' ', '_');

            // Append the column details
            sb.AppendFormat("{0} {1}, ", fieldName, fieldType);

            // Remove the last comma
            sb.Remove(sb.ToString().LastIndexOf(','), 1);

            // Set the command text
            retCommand.CommandText = sb.ToString();

            // Return the command
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