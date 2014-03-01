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

        protected override string TableName
        {
            get { return SprConstants.MdbTagTable; }
        }
    }
}
