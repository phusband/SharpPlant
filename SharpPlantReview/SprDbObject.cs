//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Data;

namespace SharpPlant.SharpPlantReview
{
    public abstract class SprDbObject : IEquatable<SprDbObject>, IComparable<SprDbObject>, IDisposable
    {
        #region Properties

        protected bool _disposed;

        /// <summary>
        ///     The backing DataRow that contains the SprDbObject values.
        /// </summary>
        internal DataRow Row
        {
            get { return _row ?? (_row = GetDataRow()); }
            set { _row = value; }
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
            Application = SprApplication.ActiveApplication;
            _row = DefaultRow;
            
        }
        protected SprDbObject(DataRow datarow)
        {
            Application = SprApplication.ActiveApplication;
            _row = datarow;
        }
        ~SprDbObject()
        {
            Dispose(false);
        }

        #endregion

        #region Methods

        protected abstract DataRow GetDataRow();
        protected abstract DataRow GetDefaultRow();

        /// <summary>
        ///     Loads the latest SprDbObject information from the MDB database.
        /// </summary>
        public void Refresh()
        {
            Row = Application.RefreshRow(TableName, Id);

            if (Row.RowState == DataRowState.Detached)
                Row.Table.Rows.Add(Row);

            Row.AcceptChanges();
        }

        /// <summary>
        ///     Updates the MDB database with the current SprDbObject information.
        /// </summary>
        public void Update()
        {
            Application.UpdateRow(TableName, Id);
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

        #region IComparable

        public int CompareTo(SprDbObject other)
        {
            return this.Id.CompareTo(other.Id);
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                    Row = null;

                _disposed = true;
            }
        }

        #endregion
    }
}
