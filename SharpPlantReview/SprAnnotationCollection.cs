//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Windows.Forms;

namespace SharpPlant.SharpPlantReview
{
    public class SprAnnotationCollection : SprDbObjectCollection<SprAnnotation>
    {
        #region Properties

        public override string TableName
        {
            get { return "text_annotations"; }
        }

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

        internal SprAnnotationCollection() : this(SprApplication.ActiveApplication) { }
        internal SprAnnotationCollection(SprApplication application) : base(application)
        {
            SetVisibility(SprTagVisibility.None);
        }

        #endregion

        #region Methods

        /// <summary>
        ///     Sets the application annotation visibility state.
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

            // Set the annotation visibility
            SendKeys.SendWait(string.Format("%GS{0}", alias));
        }

        #endregion
    }
}
