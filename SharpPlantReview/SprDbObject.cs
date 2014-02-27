//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Data;

namespace SharpPlant.SharpPlantReview
{
    public abstract class SprDbObject : IEquatable<SprDbObject>
    {
        #region Properties

        /// <summary>
        ///     The backing DataRow that contains the SprDbObject values.
        /// </summary>
        internal DataRow Row
        {
            get { return _row ?? (_row = GetDataRow()); }
        }
        private DataRow _row;

        /// <summary>
        ///     The parent SprApplication object.
        /// </summary>
        public SprApplication Application { get; private set; }

        /// <summary>
        ///     The default DataRow for instanciating a new SprDbObject.
        /// </summary>
        public DataRow DefaultRow
        {
            get { return _defaultRow ?? (_defaultRow = GetDefaultRow()); }
            set { _defaultRow = value; }
        }
        private DataRow _defaultRow;

        /// <summary>
        ///     The SprDbObject unique identification number.
        /// </summary>
        public int Id
        {
            get { return Convert.ToInt32(Row[IdKey]); }
            internal set { Row[IdKey] = value; }
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
            _row = DefaultRow;
            Application = SprApplication.ActiveApplication;
        }
        protected SprDbObject(DataRow datarow)
        {
            _row = datarow;
            Application = SprApplication.ActiveApplication;
        }

        #endregion

        #region Methods

        // Getters
        private DataRow GetDataRow()
        {
            var dataRow = Application.MdbDatabase.Tables[TableName].Rows.Find(Id);
            return dataRow ?? (DefaultRow);
        }
        protected abstract DataRow GetDefaultRow();

        /// <summary>
        ///     Loads the latest SprDbObject information from the MDB database.
        /// </summary>
        public void Refresh()
        {
            var updatedTable = DbMethods.GetDbTable(Application.MdbPath, TableName);
            var updatedRow = updatedTable.Rows.Find(Id);

            _row = updatedRow;

            if (_row.RowState == DataRowState.Detached)
                _row.Table.Rows.Add(_row);

            _row.AcceptChanges();
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

        #region IEquatable

        public bool Equals(SprDbObject other)
        { 
            if (ReferenceEquals(other, null))
                return false;
            
            if (ReferenceEquals(this, null))
                return true;               

            // Matching Ids will only work from the same table
            return Id == other.Id && other.TableName.Equals(TableName);
        }

        public override bool Equals(Object obj)
        {
            return Equals(obj as SprDbObject);
        }

        public override int GetHashCode()
        {
            return Equals(null)
                ? 0
                : new {Id, TableName}.GetHashCode();
        }

        #endregion
    }
}
