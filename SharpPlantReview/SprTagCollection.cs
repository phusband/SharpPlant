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

        public override string TableName { get { return "tag_data"; } }

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
        internal SprTagCollection(SprApplication application) : base(application)
        {
            SetVisibility(SprTagVisibility.None);
        }

        #endregion

        #region Methods

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
