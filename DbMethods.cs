//
//  Copyright © 2013 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Data;
using System.Data.OleDb;
using System.Text;

// OleDB Database methods

namespace SharpPlant
{
    internal static class DbMethods
    {
        public static DataTable GetDbTable(string dbPath, string tableName)
        {
            // Create the return table
            var returnTable = new DataTable(tableName);

            // Create the connection
            using (var connection = GetConnection(dbPath))
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

                // Return nothing on error
                catch (OleDbException)
                {
                    return null;
                }
            }

            // Return the table
            return returnTable;
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
                catch (OleDbException)
                {
                    return false;
                }
            }

            // Return true on success
            return true;
        }

        public static bool AddDbField(string dbPath, string fieldName)
        {
            // Create the connection
            using (var connection = GetConnection(dbPath))
            {
                try
                {
                    // Open the connection
                    connection.Open();

                    // Create the update command
                    var updateCommand = GetAddFieldCommand("tag_data", fieldName, connection);

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
            // Build the connection string
            var connectionString = string.Format("Provider=Microsoft.Ace.OLEDB.12.0;Data Source ={0};", dbPath);

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
                        SourceColumn = col.ColumnName
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

        private static OleDbCommand GetAddFieldCommand(string tableName, string fieldName, OleDbConnection connection)
        {
            // Create the return command
            var retCommand = connection.CreateCommand();

            // Build the command string
            var sb = new StringBuilder(string.Format("ALTER TABLE {0} ADD COLUMN ", tableName));

            // Replace any spaces from the field name
            fieldName = fieldName.Replace(' ', '_');

            // Append the column details
            sb.AppendFormat("{0} TEXT(255), ", fieldName);

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
                default:
                    return OleDbType.Variant;
            }
        }
    }
}