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

        public TObject this[int index]
        {
            get { return InnerCollection[index]; }
            set { InnerCollection[index] = value; }
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
                if (value == null) throw new ArgumentNullException("value");
                this[id] = value;
            }
        }

        #endregion

        #region Constructors

        protected SprDbObjectCollection() : this(SprApplication.ActiveApplication) { }
        protected SprDbObjectCollection(SprApplication application)
        {
            Application = application;
            InnerCollection = new List<TObject>();
            _isReadOnly = false;
            Refresh();
        }

        #endregion

        #region Methods

        protected abstract DataTable GetTable();

        protected void Refresh()
        {
            Clear();
            var updatedTable = DbMethods.GetDbTable(Application.MdbPath, Table.TableName);

            if (updatedTable == null)
                throw new SprException("Could not access MDB table {0}.", Table.TableName);

            foreach (DataRow objRow in updatedTable.Rows)
            {
                var dbObj = new TObject();
                dbObj.Row.ItemArray = objRow.ItemArray;
                InnerCollection.Add(dbObj);
            }
        }
        protected void Update()
        {
            if (!DbMethods.UpdateDbTable(Application.MdbPath, Table))
                throw new SprException("Error updating Mdb table {0}", Table.TableName);
        }

        /// <summary>
        ///     Adds a SprDbObject as a new row in the Mdb table.
        /// </summary>
        public virtual void Add(TObject item)
        {
            if (!Contains(item.Id))
            {
                InnerCollection.Add(item);
                var newRow = Table.NewRow();
                newRow.ItemArray = item.Row.ItemArray;

                Update();
            }
            else
                throw new SprException("A {0} with Id: {1} already exists in {2}",
                                       item.GetType().ToString(), item.Id, Application.MdbPath);
        }

        /// <summary>
        ///     Creates a new data field in the Mdb table.
        /// </summary>
        /// <param name="fieldName">The string name of the field to be added.  Spaces in the field name are replaced.</param>
        /// <param name="typeName"></param>
        public virtual void AddDataField(string fieldName, string typeName = "TEXT(255)")
        {
            DbMethods.AddDbField(Application.MdbPath, Table.TableName, typeName, typeName);

            Application.MdbDatabase = null;
            Table = null;
            Refresh();
            
            // Add the tag field to the MDB database
            //return DbMethods.AddDbField(MdbPath, "tag_data", fieldName);
        }

        /// <summary>
        ///     Clears all the SprDbObjects from the MDB table.
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
            return Enumerable.Contains(InnerCollection, item);
        }

        /// <summary>
        ///     Determines whether a SprDbObject of a specific Id is in the collection.
        /// </summary>
        /// <param name="id">The Id of the SprDbObject to search for.</param>
        public virtual bool Contains(int id)
        {
            return InnerCollection.Any(obj => obj.Id == id);
        }

        /// <summary>
        ///     Copies the entire SprDbObject collection to a one-dimensional array.
        /// </summary>
        public void CopyTo(TObject[] array, int arrayIndex = 0)
        {
            for (var i = arrayIndex; i < InnerCollection.Count; i++)
                array[i] = InnerCollection[i];
        }

        /// <summary>
        ///     An IEnumerator object that can be used to iterate through the collection.
        /// </summary>
        public IEnumerator<TObject> GetEnumerator()
        {
            return InnerCollection.GetEnumerator();
            //using (var iter = innerCollection.GetEnumerator())
            //{
            //    while (iter.MoveNext())
            //        yield return iter.Current;
            //}
        }
        
        /// <summary>
        ///     Removes a specific SprDbObject from the collection.
        /// </summary>
        /// <param name="item">The SprDbObject to remove from the list.</param>
        public virtual bool Remove(TObject item)
        {
            for (var i = 0; i < InnerCollection.Count; i++)
            {
                var curObj = InnerCollection[i];
                if (!item.Equals(curObj))
                    continue;

                InnerCollection.RemoveAt(i);
                Table.Rows.Find(item.Id).Delete();
                return true;
            }

            return false;
        }

        /// <summary>
        ///     Removes a SprDbObject with a matching Id from the collection.
        /// </summary>
        /// <param name="id">The Id to remove.</param>
        public bool RemoveId(int id)
        {
            return Remove(this[id.ToString()]);
        }

        #endregion
    }
}
