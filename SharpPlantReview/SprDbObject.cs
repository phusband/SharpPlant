//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Data;

namespace SharpPlant.SharpPlantReview
{
    public abstract class SprDbObject
    {
        #region Properties

        /// <summary>
        ///     The backing DataRow that contains the SprDbObject values.
        /// </summary>
        internal DataRow Row
        {
            get { return row ?? (row = GetDataRow()); }
        }
        private DataRow row;

        /// <summary>
        ///     The parent SprApplication object.
        /// </summary>
        public SprApplication Application { get; private set; }

        /// <summary>
        ///     The default DataRow for instanciating a new SprDbObject.
        /// </summary>
        public DataRow DefaultRow
        {
            get { return defaultRow ?? (defaultRow = GetDefaultRow()); }
            set { defaultRow = value; }
        }
        private DataRow defaultRow;

        /// <summary>
        ///     The SprDbObject unique identification number.
        /// </summary>
        public int Id
        {
            get { return Convert.ToInt32(Row[IdKey]); }
            private set { Row[IdKey] = value; }
        }

        /// <summary>
        ///     The unique key name for the SprDbObject Mdb table.
        /// </summary>
        public string IdKey
        { 
            get { return Row.Table.PrimaryKey[0].ColumnName; }
        }

        /// <summary>
        ///     The name of the SprDbObject Mdb table name.
        /// </summary>
        public string TableName
        {
            get { return Row.Table.TableName; }
        }

        #endregion

        #region Constructors

        protected SprDbObject()
        {
            Application = SprApplication.ActiveApplication;
            row = DefaultRow;
        }
        protected SprDbObject(DataRow datarow)
        {
            Application = SprApplication.ActiveApplication;
            row = datarow;
        }

        #endregion

        #region Methods

        // Getters
        private DataRow GetDataRow()
        {
            var dataRow = Application.MdbDatabase.Tables[TableName].Rows.Find((object)Id);
            return dataRow ?? (DefaultRow);
        }
        protected abstract DataRow GetDefaultRow();

        /// <summary>
        ///     Loads the latest SprDbObject information from the MDB database.
        /// </summary>
        public void Refresh()
        {
            var updatedTable = DbMethods.GetDbTable(Application.MdbPath, TableName);
            var updatedRow = updatedTable.Rows.Find((object)Id);

            row.ItemArray = updatedRow.ItemArray;
            row.AcceptChanges();
        }

        /// <summary>
        ///     Updates the MDB database with the current SprDbObject information.
        /// </summary>
        public void Update()
        {
            var filter = string.Format("{0} = {1}", IdKey, Id);
            DbMethods.UpdateDbTable(Application.MdbPath, filter, Row.Table);
        }

        #endregion
    }
}
