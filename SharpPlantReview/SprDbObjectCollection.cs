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
    public abstract class SprDbObjectCollection<T> : ICollection<T> where T : SprDbObject, new()
    {
        #region Properties

        internal List<T> innerCollection;

        public SprApplication Application { get; private set; }

        public int Count
        {
            get { return innerCollection.Count; }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public bool IsReadOnly
        {
            get { return isReadOnly; }
        }
        private bool isReadOnly = false;
        
        internal DataTable Table
        {
            get
            {
                return table ?? (table = Application.MdbDatabase.Tables[TableName]);
            }
            set { table = value; }
        }
        private DataTable table;

        public abstract string TableName { get; }

        public T this[int index]
        {
            get { return (T)innerCollection[index]; }
            set { innerCollection[index] = value; }
        }

        /// <summary>
        ///     Gets or sets elements at the specified SprDbObject Id.
        /// </summary>
        /// <param name="Id">The string Id of the SprDbObject to get or set.</param>
        public T this[string Id]
        {
            get { return (T)innerCollection.First(o => o.Id.ToString() == Id); }
            set
            {
                var tag = (T)innerCollection.First(o => o.Id.ToString() == Id);
                tag = value;
            }
        }

        #endregion

        #region Constructors

        protected SprDbObjectCollection() : this(SprApplication.ActiveApplication) { }
        protected SprDbObjectCollection(SprApplication application)
        {
            Application = application;
            innerCollection = new List<T>();

            Refresh();
        }

        #endregion

        #region Methods

        protected void Refresh()
        {
            Clear();
            var updatedTable = DbMethods.GetDbTable(Application.MdbPath, TableName);

            if (updatedTable == null)
                throw new SprException("Could not access MDB table {0}.", TableName);

            foreach (DataRow objRow in updatedTable.Rows)
            {
                var dbObj = new T();
                dbObj.Row.ItemArray = objRow.ItemArray;
                innerCollection.Add(dbObj);
            }
        }
        protected void Update()
        {
            if (!DbMethods.UpdateDbTable(Application.MdbPath, Table))
                throw new SprException("Error updating Mdb table {0}", TableName);
        }

        /// <summary>
        ///     Adds a SprDbObject as a new row in the Mdb table.
        /// </summary>
        /// <param name="tag">The SprDbObject to be written to the database.</param>
        public virtual void Add(T item)
        {
            if (!Contains(item.Id))
            {
                innerCollection.Add(item);
                var newRow = Table.NewRow();
                newRow.ItemArray = item.Row.ItemArray;

                Update();
            }
            else
                throw new SprException("A {0} with Id: {1} already exists in {2}",
                                       item.GetType().ToString(), item.Id, Application.MdbPath);
        }

        /// <summary>
        ///     Clears all the SprDbObjects from the MDB table.
        /// </summary>
        public virtual void Clear()
        {
            innerCollection.Clear();
            Table.Rows.Clear();
            Update();
        }

        /// <summary>
        ///     Determines whether a SprDbObject is in the collection.
        /// </summary>
        /// <param name="item">The SprDbObject to search for.</param>
        public virtual bool Contains(T item)
        {
            foreach (T dbObj in innerCollection)
            {
                if (dbObj.Equals(item))
                    return true;
            }

            return false;
        }

        /// <summary>
        ///     Determines whether a SprDbObject of a specific Id is in the collection.
        /// </summary>
        /// <param name="id">The Id of the SprDbObject to search for.</param>
        public virtual bool Contains(int id)
        {
            foreach (T obj in innerCollection)
                if (obj.Id == id)
                    return true;

            return false;
        }

        /// <summary>
        ///     Copies the entire SprDbObject collection to a one-dimensional array.
        /// </summary>
        public void CopyTo(T[] array, int arrayIndex = 0)
        {
            for (int i = arrayIndex; i < innerCollection.Count; i++)
                array[i] = (T)innerCollection[i];
        }

        /// <summary>
        ///     An IEnumerator object that can be used to iterate through the collection.
        /// </summary>
        public IEnumerator<T> GetEnumerator()
        {
            return innerCollection.GetEnumerator();
        }
        
        /// <summary>
        ///     Removes a specific SprDbObject from the collection.
        /// </summary>
        /// <param name="item">The SprDbObject to remove from the list.</param>
        public virtual bool Remove(T item)
        {
            for (int i = 0; i < innerCollection.Count; i++)
            {
                var curObj = (T)innerCollection[i];
                if (item.Equals(curObj))
                {
                    innerCollection.RemoveAt(i);
                    Table.Rows.Find(item.Id).Delete();
                    return true;
                }
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
