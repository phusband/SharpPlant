//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace SharpPlant.SharpPlantReview
{
    public abstract class SprDbObjectCollection<TObject> : ICollection<TObject> where TObject : SprDbObject, new()
    {
        #region Properties

        // Holding these in a collection should be quicker over time
        // versus de-serializing DataRows into tags on demand.
        internal List<TObject> InnerCollection;

        public SprApplication Application { get; private set; }

        public int Count
        {
            get { return InnerCollection.Count; }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public bool IsReadOnly
        {
            get { return _isReadOnly; }
        }
        private readonly bool _isReadOnly;

        internal DataTable Table
        {
            get { return _table ?? (_table = GetTable()); }
            set { _table = value; }
        }
        private DataTable _table;

        /// <summary>
        ///     Gets or sets elements at the specified index.
        /// </summary>
        /// <param name="index">The index of the SprDbObject to get or set.</param>
        public TObject this[int index]
        {
            get { return InnerCollection[index]; }
            set
            {
                if (value == null)
                    throw new ArgumentNullException("value");

                InnerCollection[index] = value;
            }
        }

        /// <summary>
        ///     Gets or sets elements at the specified SprDbObject Id.
        /// </summary>
        /// <param name="id">The string Id of the SprDbObject to get or set.</param>
        public TObject this[string id]
        {
            get { return InnerCollection.First(o => o.Id.ToString() == id); }
            set
            {
                if (value == null)
                    throw new ArgumentNullException("value");

                this[id] = value;
                //obj = value;
            }
        }

        #endregion

        #region Constructors

        protected SprDbObjectCollection() : this(SprApplication.ActiveApplication) { }
        protected SprDbObjectCollection(SprApplication application)
        {
            Application = application;
            _isReadOnly = false;
            Refresh();
        }

        #endregion

        #region Methods

        // Now the type-specific table can be defined in the specific collection
        //protected abstract DataTable GetTable();
        protected abstract string TableName { get; }
        private DataTable GetTable()
        {
            var returnTable = Application.MdbDatabase.Tables[TableName]
                ?? Application.RefreshTable(TableName);

            return returnTable;
        }

        /// <summary>
        ///     Adds the specified SprDbObject to the collection.
        /// </summary>
        public virtual void Add(TObject item)
        {
            Add(new[] { item });
        }

        /// <summary>
        ///     Adds several specified SprDbObjects to the collection. 
        /// </summary>
        /// <param name="items"></param>
        public virtual void Add(IEnumerable<TObject> items)
        {
            foreach(var item in items)
            {
                if (!Contains(item.Id))
                {
                    var newRow = Table.NewRow();
                    newRow = item.Row;

                    if (newRow.RowState == DataRowState.Detached)
                        Table.Rows.Add(newRow);

                    newRow.AcceptChanges();
                    InnerCollection.Add(item);
                }

                else
                    throw new SprException("A {0} with Id: {1} already exists in {2}",
                                           item.GetType().ToString(), item.Id, Application.MdbPath);
            }

            Update();
        }

        /// <summary>
        ///     Creates a new data field in the Mdb table.
        ///     Any pending changes to the collection will be lost.
        /// </summary>
        /// <param name="fieldName">The string name of the field to be added.  Spaces in the field name are replaced.</param>
        /// <param name="typeName">The data type of the field to be added.  Default is TEXT(255).</param>
        public virtual void AddDataField(string fieldName, string typeName = "TEXT(255)")
        {
            if (Table.Columns.Contains(fieldName))
                return;

            DbMethods.AddDbField(Application.MdbPath, Table.TableName, fieldName, typeName);

            Application.MdbDatabase = null;
            Refresh();
        }

        /// <summary>
        ///     Clears the SprDbObject collection.
        /// </summary>
        public virtual void Clear()
        {
            InnerCollection.Clear();
            Table.Rows.Clear();
            Update();
        }

        /// <summary>
        ///     Determines whether a SprDbObject is in the collection.
        /// </summary>
        /// <param name="item">The SprDbObject to search for.</param>
        public virtual bool Contains(TObject item)
        {
            return Contains(item.Id);// We'll check more later
        }

        /// <summary>
        ///     Determines whether a SprDbObject of a specific Id is in the collection.
        /// </summary>
        /// <param name="id">The Id of the SprDbObject to search for.</param>
        public virtual bool Contains(int id)
        {
            return Table.Rows.Contains(id);
            //return InnerCollection.Any(obj => obj.Id == id);
        }

        /// <summary>
        ///     Copies the entire SprDbObject collection to a one-dimensional array.
        /// </summary>
        public void CopyTo(TObject[] array, int arrayIndex = 0)
        {
            InnerCollection.CopyTo(array, arrayIndex);
        }

        /// <summary>
        ///     An IEnumerator object that can be used to iterate through the collection.
        /// </summary>
        public IEnumerator<TObject> GetEnumerator()
        {
            return InnerCollection.GetEnumerator();
        }

        /// <summary>
        ///     Resets the collection to the current MDB table values.
        /// </summary>
        public virtual void Refresh()
        {
            Table = Application.RefreshTable(TableName);
            InnerCollection = new List<TObject>();

            foreach (DataRow row in Table.Rows)
            {
                var obj = new TObject {Row = row};
                InnerCollection.Add(obj);
                if (obj.Row.RowState == DataRowState.Detached)
                    Table.Rows.Add(obj.Row);
            }

            Table.AcceptChanges();
        }
        
        /// <summary>
        ///     Removes a specific SprDbObject from the collection.
        /// </summary>
        /// <param name="item">The SprDbObject to remove from the collection.</param>
        public virtual bool Remove(TObject item)
        {
            return Remove(new[] { item });
        }

        /// <summary>
        ///     Removes several specified SprDbObjects from the collection.
        /// </summary>
        /// <param name="items">The items to remove from the collection.</param>
        public virtual bool Remove(IEnumerable<TObject> items)
        {
            foreach (var item in items)
            {
                InnerCollection.Remove(item);
                var curItem = Table.Rows.Find(item.Id);
                curItem.Delete();
            }

            Update();
            return true;
        }

        /// <summary>
        ///     Removes a SprDbObject with a matching Id from the collection.
        /// </summary>
        /// <param name="id">The Id to remove.</param>
        public bool RemoveId(int id)
        {
            return Remove(this[id.ToString()]);
        }

        /// <summary>
        ///     Updates the Mdb table with the current collection items.
        /// </summary>
        public void Update()
        {
            Application.UpdateTable(TableName);
        }

        public DataTable ToTable()
        {
            return Table.Copy();
        }

        #endregion
    }
}
