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

        public static DataTable GetDbTable(string dbPath, string tableName, string[] fields = null)
        {
            if (fields == null) fields = new string[] { "*" };

            var tbl_Return = new DataTable(tableName);
            using (var connection = GetConnection(dbPath))
            {
                try
                {
                    connection.Open();
                    var selectCommand = GetSelectCommand(connection, tableName, fields);
                    selectCommand.CommandType = CommandType.TableDirect;

                    var adapter = new OleDbDataAdapter(selectCommand);
                    adapter.FillSchema(tbl_Return, SchemaType.Mapped);
                    adapter.Fill(tbl_Return);
                }
                catch (Exception)
                {
                    
                    throw;
                }
            }

            return new DataTable();
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
            }
        }

        public static bool UpdateDbTable(string dbPath, DataRow inputRow)
        {
            return UpdateDbTable(dbPath, new DataRow[] { inputRow });
        }
        public static bool UpdateDbTable(string dbPath, DataRow[] inputRows)
        {
            // Create the connection
            using (var connection = GetConnection(dbPath))
            {
                try
                {
                    // Open the connection
                    connection.Open();

                    // Create the adapter
                    var tableName = inputRows[0].Table.TableName;
                    var dataAdapter = new OleDbDataAdapter(string.Format("SELECT * FROM {0}", tableName), connection);

                    // Create the update command
                    var commandBuilder = new OleDbCommandBuilder(dataAdapter);
                    commandBuilder.SetAllValues = false;
                    dataAdapter.UpdateCommand = commandBuilder.GetUpdateCommand(true);

                    // Update the database
                    dataAdapter.Update(inputRows);

                } // Return false on error
                catch (OleDbException)
                {
                    //throw ex;
                    return false;
                }
            }

            // Return true on success
            return true;
        }
        public static bool UpdateDbTable(string dbPath, DataTable inputTable)
        {
            // Create the connection
            using (var connection = GetConnection(dbPath))
            {
                try
                {
                    // Open the connection
                    connection.Open();

                    // Create the adapter
                    var dataAdapter = new OleDbDataAdapter(string.Format("SELECT * FROM {0}", inputTable.TableName), connection);

                    // Create the update command
                    dataAdapter.UpdateCommand = new OleDbCommandBuilder(dataAdapter).GetUpdateCommand(true);

                    // Update the database
                    dataAdapter.Update(inputTable);

                } // Return false on error
                catch (OleDbException)
                {
                    //throw ex;
                    return false;
                }
            }

            // Return true on success
            return true;
        }
        public static bool AddDbField(string dbPath, string tableName, string fieldName, string fieldType = "TEXT(255)")
        {
            // Create the connection
            using (var connection = GetConnection(dbPath))
            {
                try
                {
                    // Open the connection
                    connection.Open();

                    // Create the update command
                    var updateCommand = GetAddFieldCommand(connection, tableName, fieldName, fieldType);

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

        private static OleDbConnection GetConnection(string dbPath, OleDbDriver type = OleDbDriver.Ace)
        {
            string connectionString;

            if (type == OleDbDriver.Ace)
                connectionString = string.Format("Provider=Microsoft.Ace.OLEDB.12.0;Data Source ={0};", dbPath);
            else
                connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source ={0};", dbPath);

            return new OleDbConnection(connectionString);
        }
        private static OleDbCommand GetSelectCommand(OleDbConnection connection, string tableName, string[] fields)
        {
            var retCommand = connection.CreateCommand();
            var sb = new StringBuilder("SELECT ");

            for (int i = 0; i < fields.Length; i++)
                sb.AppendFormat("{0}, ", fields[i]);

            sb.Remove(sb.ToString().LastIndexOf(','), 1);
            sb.AppendFormat("FROM {0}", tableName);

            retCommand.CommandText = sb.ToString();
            return retCommand;
        }
        private static OleDbCommand GetAddFieldCommand(OleDbConnection connection, string tableName, string fieldName, string fieldType)
        {
            var retCommand = connection.CreateCommand();
            var sb = new StringBuilder(string.Format("ALTER TABLE {0} ADD COLUMN ", tableName));

            // Replace any spaces from the field name
            fieldName = fieldName.Replace(' ', '_');

            sb.AppendFormat("{0} {1}, ", fieldName, fieldType);
            sb.Remove(sb.ToString().LastIndexOf(','), 1);

            retCommand.CommandText = sb.ToString();
            return retCommand;
        }
        public enum OleDbDriver
        {
            Ace,
            Jet
        }
    }
}