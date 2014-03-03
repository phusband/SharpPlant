//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;

namespace SharpPlant.SharpPlantReview
{
    public class SprTagCollection : SprDbObjectCollection<SprTag>
    {
        #region Properties

        public SprTagVisibility Visibility
        {
            get { return _visibility; }
            set 
            {
                SetVisibility(value);
                _visibility = value; 
            }
        }
        private SprTagVisibility _visibility;

        #endregion

        #region Constructors

        internal SprTagCollection() : this(SprApplication.ActiveApplication) { }
        internal SprTagCollection(SprApplication application) : base(application)
        {
        }

        #endregion

        #region Methods

        protected override string TableName
        {
            get { return SprConstants.MdbTagTable; }
        }

        public override void Refresh()
        {
            Table = Application.RefreshTable(TableName);
            InnerCollection = new List<SprTag>();

            foreach (DataRow row in Table.Rows)
            {
                var tag = new SprTag { Row = row };
                InnerCollection.Add(tag);
                tag.Collection = this;

                if (tag.Row.RowState == DataRowState.Detached)
                    Table.Rows.Add(tag.Row);
            }

            Table.AcceptChanges();
        }

        public override bool Remove(SprTag item)
        {
            InnerCollection.Remove(item);
            var curItem = Table.Rows.Find(item.Id);
            Table.Rows.Remove(curItem);

            Update();
            return true;
        }

        /// <summary>
        ///     Sets the application tag visibility state.
        /// </summary>
        /// <param name="state"></param>
        public void SetVisibility(SprTagVisibility state)
        {
            if (!Application.IsConnected)
                throw SprExceptions.SprNotConnected;

            Application.Windows.TextWindow.Clear();
            Application.Activate();

            // Get the menu alias character from the enumerator
            var alias = Char.ConvertFromUtf32((int)state);

            // Set the tag visibility
            SendKeys.SendWait(string.Format("%GS{0}", alias));
        }

        #endregion
    }
}
