//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace SharpPlant.SharpPlantReview
{
    public class SprTagCollection : ICollection<SprTag>
    {
        #region Properties

        internal class SprTagEnumerator : IEnumerator<SprTag>
        {
            private SprTagCollection _collection;
            private int _curIndex;
            private SprTag _curTag;

            public SprTag Current
            {
                get { return _curTag; }
            }
            object IEnumerator.Current
            {
                get { return Current; }
            }

            internal SprTagEnumerator(SprTagCollection collection)
            {
                _collection = collection;
                _curIndex = -1;
                _curTag = default(SprTag);
            }

            public void Dispose() { }
            public bool MoveNext()
            {
                if (++_curIndex >= _collection.Count)
                    return false;
                else
                    _curTag = _collection[_curIndex];
                return true;
            }
            public void Reset()
            {
                _curIndex = -1;
            }
        }

        /// <summary>
        ///     Gets or sets elements at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index of the SprTag to get or set.</param>
        public SprTag this[int index]
        {
            get { return (SprTag)innerCollection[index]; }
            set { innerCollection[index] = value; }
        }

        /// <summary>
        ///     Gets or sets elements at the specified TagId.
        /// </summary>
        /// <param name="tagId">The string Id of the SprTag to get or set.</param>
        public SprTag this[string tagId]
        {
            get { return (SprTag)innerCollection.First(t => t.Id.ToString() == tagId); }
            set
            {
                var tag = (SprTag)innerCollection.First(t => t.Id.ToString() == tagId);
                tag = value;
            }
        }

        private List<SprTag> innerCollection;

        public bool IsReadOnly
        {
            get { return isReadOnly; }
        }
        private bool isReadOnly = false;

        public IEnumerator<SprTag> GetEnumerator()
        {
            return new SprTagEnumerator(this);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return new SprTagEnumerator(this);
        }

        public SprApplication Application { get; private set; }

        public int Count
        {
            get { return innerCollection.Count; }
        }

        internal DataTable Table
        {
            get 
            {
                if (table == null)
                    table = Application.MdbDatabase.Tables["tag_data"];
                return table;
            }
            set { table = value; }
        }
        private DataTable table;

        public SprTagVisibility Visibility
        {
            get { return visibility; }
            set 
            {
                SetVisibility(value);
                visibility = value; 
            }
        }
        private SprTagVisibility visibility;

        #endregion

        #region Constructors

        internal SprTagCollection() : this(SprApplication.ActiveApplication) { }
        internal SprTagCollection(SprApplication application)
        {
            Application = application;
            innerCollection = new List<SprTag>();

            if (Table == null)
                throw new SprException("Could not access tag table.");

            foreach (DataRow tagRow in Table.Rows)
                innerCollection.Add(new SprTag(tagRow));

            SetVisibility(SprTagVisibility.None);
        }

        #endregion

        #region Methods

        private void Refresh()
        {
            Clear();
            var updatedTable = DbMethods.GetDbTable(Application.MdbPath, Table.TableName);

            foreach (DataRow tagRow in updatedTable.Rows)
                innerCollection.Add(new SprTag(tagRow));
        }
        private void Update()
        {
            DbMethods.UpdateDbTable(Application.MdbPath, Table);
        }
        private void SetVisibility(SprTagVisibility state)
        {
            if (!Application.IsConnected)
                throw SprExceptions.SprNotConnected;

            Application.TextWindow_Clear();
            Application.Activate();

            // Get the menu alias character from the enumerator
            var alias = Char.ConvertFromUtf32((int)state);

            // Set the tag visibility
            SendKeys.SendWait(string.Format("%GS{0}", alias));
        }

        /// <summary>
        ///     Creates a new data field in the Mdb tag_data table.
        ///     Returns true if the field already exists.
        /// </summary>
        /// <param name="fieldName">The string name of the field to be added.  Spaces in the field name are replaced.</param>
        /// <returns>Indicates the success or failure of the table modification.</returns>
        //public bool Tags_AddDataField(string fieldName)
        //{
        //    // Throw an exception if not connected
        //    if (!IsConnected) throw SprExceptions.SprNotConnected;

        //    // Add the tag field to the MDB database
        //    return DbMethods.AddDbField(MdbPath, "tag_data", fieldName);
        //}

        /// <summary>
        ///     Adds a tag as a new row in the Mdb tag_data table.
        /// </summary>
        /// <param name="tag">The Tag to be written to the database.</param>
        public void Add(SprTag item)
        {
            if (!Contains(item.Id))
            {
                innerCollection.Add(item);
                var newRow = Table.NewRow();
                newRow.ItemArray = item.Row.ItemArray;

                Update();
            }
            else
                throw new SprException(string.Format("A Tag with Id: {0} already exists in {1}",
                                                     item.Id, Application.MdbPath));
        }

        /// <summary>
        ///     Clears all the tags from the MDB database.
        /// </summary>
        public void Clear()
        {
            innerCollection.Clear();
            Table.Rows.Clear();
            Update();
        }

        /// <summary>
        ///     Determines whether a tag is in the collection.
        /// </summary>
        /// <param name="item">The SprTag to search for.</param>
        public bool Contains(SprTag item)
        {
            foreach (SprTag tag in innerCollection)
            {
                if (tag.Equals(item))
                    return true;
            }

            return false;
        }

        /// <summary>
        ///     Determines whether a tag of a specific Id is in the collection.
        /// </summary>
        /// <param name="id">The Id of the SprTag to search for.</param>
        public bool Contains(int id)
        {
            foreach (SprTag tag in innerCollection)
            {
                if (tag.Id == id)
                    return true;
            }

            return false;
        }

        /// <summary>
        ///     Copies the entire SprTag list to a one-dimensional array.
        /// </summary>
        public void CopyTo(SprTag[] array, int arrayIndex = 0)
        {
            for (int i = arrayIndex; i < innerCollection.Count; i++)
                array[i] = (SprTag)innerCollection[i];
        }

        /// <summary>
        ///     Removes a specific SprTag from the collection.
        /// </summary>
        /// <param name="item">The SprTag to remove from the list.</param>
        public bool Remove(SprTag item)
        {
            for (int i = 0; i < innerCollection.Count; i++)
            {
                var curTag = (SprTag)innerCollection[i];
                if (item.Equals(curTag))
                {
                    innerCollection.RemoveAt(i);
                    Table.Rows.Find(item.Id).Delete();
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        ///     Removes a tag with a matching Id from the collection.
        /// </summary>
        /// <param name="id">The Id to match the tag to remove.</param>
        public bool RemoveId(int id)
        {
            return Remove(this[id.ToString()]);
        }

        #endregion
    }
}
