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
                if (text == null)
                    text = GetText();
                return text;
            }
            set
            {
                Update(Title, value);
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
                if (title == null)
                    title = GetTitle();
                return title;
            }
            set
            {
                Update(Text, value);
                title = value;
            }
        }
        private string title;

        #endregion

        #region Constructors

        internal SprTextWindow(SprApplication application) : base(application, SprWindowType.TextWindow)
        {
            text = GetText();
            title = GetTitle();
        }

        #endregion

        #region Methods

        private string GetText()
        {
            if (!Application.IsConnected)
                throw SprExceptions.SprNotConnected;

            var orgTitle = string.Empty;
            var orgText = string.Empty;
            int orgLength;

            // Get the existing text window values
            Application.SprStatus = Application.DrApi.TextWindowGet(ref orgTitle, out orgLength, ref orgText);

            // Set an empty string for null values
            return orgText ?? (string.Empty);
        }
        private string GetTitle()
        {
            if (!Application.IsConnected)
                throw SprExceptions.SprNotConnected;

            var orgTitle = string.Empty;
            var orgText = string.Empty;
            int orgLength;

            // Get the existing text window values
            Application.SprStatus = Application.DrApi.TextWindowGet(ref orgTitle, out orgLength, ref orgText);

            // Return the title, empty string if null
            return orgTitle ?? (string.Empty);
        }
        private void Update(string mainText, string titleText)
        {
            if (!Application.IsConnected)
                throw SprExceptions.SprNotConnected;

            Application.SprStatus = Application.DrApi.TextWindow(SprConstants.SprClearTextWindow, titleText, mainText, 0);
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
