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
        public abstract DataRow DefaultRow { get; }

        /// <summary>
        ///     The unique key name for the SprDbObject Mdb table.
        /// </summary>
        public abstract string PrimaryKey { get; }

        /// <summary>
        ///     The name of the SprDbObject Mdb table name.
        /// </summary>
        public string TableName { get { return Row.Table.TableName; } }

        /// <summary>
        ///     The SprDbObject unique identification number.
        /// </summary>
        public int Id
        {
            get { return Convert.ToInt32(Row[PrimaryKey]); }
            private set { Row[PrimaryKey] = value; }
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

        private DataRow GetDataRow()
        {
            var dataRow = Application.MdbDatabase.Tables[TableName].Rows.Find((object)Id);
            return dataRow ?? (DefaultRow);
        }

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
            var filter = string.Format("{0} = {1}", PrimaryKey, Id);
            DbMethods.UpdateDbTable(Application.MdbPath, filter, Row.Table);
        }

        #endregion
    }
}
