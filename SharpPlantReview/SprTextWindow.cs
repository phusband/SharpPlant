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
                return text;
            }
            set
            {
                UpdateText();
                text = value;
            }
        }
        private string text;

        /// <summary>
        ///     The title of the window.
        /// </summary>
        public string Title
        {
            get 
            {
                RefreshText();
                return title;
            }
            set
            {
                UpdateText();
                title = value;
            }
        }
        private string title;

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

            var orgTitle = title;
            var orgText = text;
            int orgLength = 0;

            try
            {
                // Get the existing text window values
                Application.SprStatus = Application.DrApi.TextWindowGet(ref orgTitle, out orgLength, ref orgText);
            }
            catch (SprException)
            {
                return;
            }
            finally
            {
                text = orgText;
                title = orgTitle;
            }
        }
        private void UpdateText()
        {
            if (!Application.IsConnected)
                throw SprExceptions.SprNotConnected;

            var flags = 0;
            flags |= SprConstants.SprClearTextWindow;

            if (title == null)
                title = "Text View";

            if (text == null)
                text = string.Empty;

            // Get the text window values
            Application.SprStatus = Application.DrApi.TextWindow(flags, title, text, 0);
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
