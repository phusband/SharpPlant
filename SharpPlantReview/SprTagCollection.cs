//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Data;
using System.Windows.Forms;

namespace SharpPlant.SharpPlantReview
{
    public class SprTagCollection : SprDbObjectCollection<SprTag>
    {
        #region Properties

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
        }

        #endregion

        #region Methods

        protected override DataTable GetTable()
        {
            return Application.MdbDatabase.Tables["tag_data"];
            //return Application.Tags.Table;
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
