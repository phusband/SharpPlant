//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
namespace SharpPlant.SharpPlantReview
{
    /// <summary>
    ///     Provides the structure containing information about a SmartPlant Review window.
    /// </summary>
    public class SprWindow
    {
        #region Properties

        /// <summary>
        ///     The active COM reference to the DrWindow class
        /// </summary>
        internal dynamic DrWindow;

        /// <summary>
        ///     The parent Application reference.
        /// </summary>
        public SprApplication Application { get; private set; }

        /// <summary>
        ///     Height of the working area of the window.
        /// </summary>
        public int Height
        {
            get
            {
                Refresh();
                return DrWindow.Height;
            }
            set
            {
                DrWindow.Height = value;
                Update();
            }
        }

        /// <summary>
        ///     Left(x) position of the top-left corner of the window.
        /// </summary>
        public int Left
        {
            get
            {
                Refresh();
                return DrWindow.Left;
            }
            set
            {
                DrWindow.Left = value;
                Update();
            }
        }

        /// <summary>
        ///     Top (y) position of the top-left corner of the window.
        /// </summary>
        public int Top
        {
            get
            {
                Refresh();
                return DrWindow.Top;
            }
            set
            {
                DrWindow.Top = value;
                Update();
            }
        }

        /// <summary>
        ///     Width of the working area of the window.
        /// </summary>
        public int Width
        {
            get
            {
                Refresh();
                return DrWindow.Width;
            }
            set
            {
                DrWindow.Width = value;
                Update();
            }
        }

        /// <summary>
        ///     The window handle.
        /// </summary>
        public int HWnd
        {
            get { return _hwnd; }
        }
        private int _hwnd;

        /// <summary>
        ///     Index of the window.
        /// </summary>
        public SprWindowType Type { get; private set; }

        #endregion

        #region Constructors

        internal SprWindow(SprApplication application, SprWindowType type)
        {
            Application = application;

            Type = type;

            // Get a new DrSnapShot object
            DrWindow = Activator.CreateInstance(SprImportedTypes.DrWindow);

            Refresh();
        }

        #endregion

        #region Methods

        private void Refresh()
        {
            // Get the window data
            Application.SprStatus = Application.DrApi.WindowGet((int)Type, out DrWindow);
            if (Application.SprStatus != 0)
                throw Application.SprException;

            Application.SprStatus = Application.DrApi.WindowHandleGet((int)Type, out _hwnd);
            if (Application.SprStatus != 0)
                throw Application.SprException;

        }
        private void Update()
        {
            // Set the window data
            Application.SprStatus = Application.DrApi.WindowSet((int)Type, DrWindow);
            if (Application.SprStatus != 0)
                throw Application.SprException;
        }

        #endregion
    }
}