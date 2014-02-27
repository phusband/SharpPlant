//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

namespace SharpPlant.SharpPlantReview
{
    /// <summary>
    ///     Provides the structure containing information about the SmartPlant Review text window.
    /// </summary>
    public class SprTextWindow : SprWindow
    {
        #region Properties

        /// <summary>
        ///     The text currently shown in the window.
        /// </summary>
        public string Text
        {
            get
            {
                RefreshText();
                return _text;
            }
            set
            {
                UpdateText();
                _text = value;
            }
        }
        private string _text;

        /// <summary>
        ///     The title of the window.
        /// </summary>
        public string Title
        {
            get 
            {
                RefreshText();
                return _title;
            }
            set
            {
                UpdateText();
                _title = value;
            }
        }
        private string _title;

        #endregion

        #region Constructors

        internal SprTextWindow(SprApplication application) : base(application, SprWindowType.TextWindow)
        {
            RefreshText();
        }

        #endregion

        #region Methods

        private void RefreshText()
        {
            if (!Application.IsConnected)
                throw SprExceptions.SprNotConnected;

            var orgTitle = _title;
            var orgText = _text;

            try
            {
                // Get the existing text window values
                int orgLength;
                Application.SprStatus = Application.DrApi.TextWindowGet(ref orgTitle, out orgLength, ref orgText);
                if (Application.SprStatus != 0)
                    throw Application.SprException;

            }
            catch (SprException)
            {
                // Do nothing, because this method crashes when no text exists
                // in the text window.  Stay classy Intergraph.
            }
            finally
            {
                _text = orgText;
                _title = orgTitle;
            }
        }
        private void UpdateText()
        {
            if (!Application.IsConnected)
                throw SprExceptions.SprNotConnected;

            var flags = 0;
            flags |= SprConstants.SprClearTextWindow;

            if (_title == null)
                _title = "Text View";

            if (_text == null)
                _text = string.Empty;

            // Get the text window values
            Application.SprStatus = Application.DrApi.TextWindow(flags, _title, _text, 0);
            if (Application.SprStatus != 0)
                throw Application.SprException;
        }

        /// <summary>
        ///     Clears the contents of the SmartPlant Review text window.
        /// </summary>
        public void Clear()
        {
            if (!Application.IsConnected)
                throw SprExceptions.SprNotConnected;

            // Send a blank string to the application text window
            Application.SprStatus = Application.DrApi.TextWindow(SprConstants.SprClearTextWindow, "Text View", string.Empty, 0);
        }

        #endregion
    }
}
