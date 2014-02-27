//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Data;
using System.Windows.Forms;

namespace SharpPlant.SharpPlantReview
{
    public class SprAnnotationCollection : SprDbObjectCollection<SprAnnotation>
    {
        #region Properties

        public bool Visibility
        {
            get { return visibility; }
            set
            {
                SetVisibility(value);
                visibility = value;
            }
        }
        private bool visibility;

        #endregion

        #region Constructors

        internal SprAnnotationCollection() : this(SprApplication.ActiveApplication) { }
        internal SprAnnotationCollection(SprApplication application) : base(application)
        {
            //SetVisibility(false);
        }

        #endregion

        #region Methods

        protected override DataTable GetTable()
        {
            return Application.MdbDatabase.Tables["text_annotations"];
            //return Application.Annotations.Table;
        }

        /// <summary>
        ///     Sets the application annotation visibility state.
        /// </summary>
        public void SetVisibility(bool visible)
        {
            if (!Application.IsConnected)
                throw SprExceptions.SprNotConnected;

            var visState = Convert.ToInt32(visible);

            // Create the view object
            dynamic objViewdataDbl = Activator.CreateInstance(SprImportedTypes.DrViewDbl);

            // Set the view object as the SPR Application main view
            Application SprStatus = Application.DrApi.ViewGetDbl(0, ref objViewdataDbl);

            // Apply the updated annotation display
            objViewdataDbl.AllAnnotationsDisplay = visState;

            // Update the global annotation visibility properties
            Application.SprStatus = Application.DrApi.GlobalOptionsSet(SprConstants.SprGlobalAnnoDisplay, visState);
            Application.SprStatus = Application.DrApi.GlobalOptionsSet(SprConstants.SprGlobalAnnoTextDisplay, visState);
            Application.SprStatus = Application.DrApi.GlobalOptionsSet(SprConstants.SprGlobalAnnoDataDisplay, visState);

            // Update the main view in SPR
            Application.SprStatus = Application.DrApi.ViewSetDbl(0, ref objViewdataDbl);
        }

        #endregion
    }
}
