using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;

namespace SharpPlant.SmartPlantReview
{
    /// <summary>
    ///     Provides the properties for creating a snapshot in SmartPlant Review.
    /// </summary>
    public class SnapShot
    {
        #region SnapShot Properties

        /// <summary>
        ///     The active COM reference to the DrSnapShot class
        /// </summary>
        internal dynamic DrSnapShot;

        // Controls the DrSnapShot flag
        internal int Flags;

        /// <summary>
        ///     The parent Application reference.
        /// </summary>
        public Application Application { get; private set; }

        /// <summary>
        ///     Anti-aliasing factor (1 to 4); 1 = no anti-aliasing.
        /// </summary>
        public int AntiAlias
        {
            get
            {
                if (IsActive) return DrSnapShot.AntiAlias;
                return -1;
            }
            set
            {
                if (value < 1) DrSnapShot.AntiAlias = 1;
                else if (AntiAlias > 4) DrSnapShot.AntiAlias = 4;
                else DrSnapShot.AntiAlias = value;
            }
        }

        /// <summary>
        ///     Pixel height of the output image (minimum 10).  Overriden by Scale/AspectOn.
        /// </summary>
        public int Height
        {
            get
            {
                if (IsActive) return DrSnapShot.Height;
                return -1;
            }
            set
            {
                if (!IsActive) return;
                DrSnapShot.Height = value > 10 ? value : 10;
            }
        }

        /// <summary>
        ///     Pixel width of the output image (minimum 10).  Overriden by Scale/AspectOn.
        /// </summary>
        public int Width
        {
            get
            {
                if (IsActive) return DrSnapShot.Width;
                return -1;
            }
            set
            {
                if (!IsActive) return;
                DrSnapShot.Width = value > 10 ? value : 10;
            }
        }

        /// <summary>
        ///     Output image size / main application view size.  Only applied if AspectOn is true.
        /// </summary>
        public double Scale
        {
            get
            {
                if (IsActive) return DrSnapShot.Scale;
                return -1;
            }
            set { if (IsActive) DrSnapShot.Scale = value; }
        }

        /// <summary>
        ///     Output format of the snapshot.
        /// </summary>
        public SnapshotFormat OutputFormat { get; set; }

        /// <summary>
        ///     Determines whether the snapshot will overwrite an existing snapshot of the same name.
        /// </summary>
        public bool Overwrite
        {
            get
            {
                // Return the bitwise zero check
                return (Flags & Constants.FILE_OVERWRITE_OK) != 0;
            }
            set
            {
                // Set flag true/false
                if (value) Flags |= Constants.FILE_OVERWRITE_OK;
                else Flags &= ~Constants.FILE_OVERWRITE_OK;
            }
        }

        /// <summary>
        ///     Determines if the snapshot is fullscreen
        /// </summary>
        public bool Fullscreen
        {
            get
            {
                // Return the bitwise zero check
                return (Flags & Constants.SNAP_FULL_SCREEN) != 0;
            }
            set
            {
                // Set flag true/false
                if (value) Flags |= Constants.SNAP_FULL_SCREEN;
                else Flags &= ~Constants.SNAP_FULL_SCREEN;
            }
        }

        /// <summary>
        ///     Determines if the snapshot is rotated 90 degrees clockwise.
        /// </summary>
        public bool Rotated90
        {
            get
            {
                // Return the bitwise zero check
                return (Flags & Constants.SNAP_ROTATE_90) != 0;
            }
            set
            {
                // Set flag true/false
                if (value) Flags |= Constants.SNAP_ROTATE_90;
                else Flags &= ~Constants.SNAP_ROTATE_90;
            }
        }

        /// <summary>
        ///     Determines whether the scale property is used to determine output image size.
        /// </summary>
        public bool AspectOn
        {
            get
            {
                // Return the bitwise zero check
                return (Flags & Constants.SNAP_ASPECT_ON) != 0;
            }
            set
            {
                // Set flag true/false
                if (value) Flags |= Constants.SNAP_ASPECT_ON;
                else Flags &= ~Constants.SNAP_ASPECT_ON;
            }
        }

        // Determines whether a reference to the COM object is established
        private bool IsActive
        {
            get { return (DrSnapShot != null); }
        }

        #endregion

        // SnapShot initializer
        public SnapShot()
        {
            // Link the parent application
            Application = SmartPlantReview.ActiveApplication;

            // Get a new DrSnapShot object
            DrSnapShot = Activator.CreateInstance(ImportedTypes.DrSnapShot);

            // Set the default settings flags
            Flags |= Constants.FILE_OVERWRITE_OK | Constants.SNAP_FORCE_BMP;

            // Set the Antialias default
            AntiAlias = 2;

            // Set the default size values
            try
            {
                // Get the application's main window size 
                dynamic drWindow = Activator.CreateInstance(ImportedTypes.DrWindow);
                Application.DrApi.WindowGet(0, drWindow);

                // Set the returned sizes
                Height = drWindow.Height;
                Width = drWindow.Width;
            }
            catch
            {
                // On error manually set the values
                Height = 800;
                Width = 1000;
            }
        }

        // Internal snapshot methods
        internal static bool FormatSnapshot(string imagePath, SnapshotFormat format)
        {
            // Get the saved image
            var curImage = Image.FromFile(imagePath);
            var result = new Bitmap(curImage.Width, curImage.Height);

            // Correct the resolution
            result.SetResolution(curImage.HorizontalResolution, curImage.VerticalResolution);

            // Draw the image from the bitmap
            using (var g = Graphics.FromImage(result))
            {
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.DrawImage(curImage, 0, 0, result.Width, result.Height);
            }

            switch (format)
            {
                // Save as PNG
                case SnapshotFormat.Png:
                    result.Save(imagePath.Replace(".bmp", ".png"), ImageFormat.Png);
                    break;

                // Save as JPG
                case SnapshotFormat.Jpg:
                    {
                        // Create a custom JPG encoder to manually set the quality
                        var jpgEncoder = GetEncoder(ImageFormat.Jpeg);
                        var encodeQuality = Encoder.Quality;
                        var encodeParams = new EncoderParameters(1);
                        var qualityParam = new EncoderParameter(encodeQuality, 95L);
                        encodeParams.Param[0] = qualityParam;

                        // Save to JPG using the custom encoder
                        result.Save(imagePath.Replace(".bmp", ".jpg"), jpgEncoder, encodeParams);
                    }
                    break;
            }

            // Delete the original image
            while (File.Exists(imagePath))
            {
                try
                {
                    curImage.Dispose();
                    result.Dispose();
                    File.Delete(imagePath);
                }
                catch (IOException)
                {
                    Thread.Sleep(100);
                }
            }

            return true;
        }

        internal static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            var codecs = ImageCodecInfo.GetImageDecoders();
            return codecs.FirstOrDefault(codec => codec.FormatID == format.Guid);
        }
    }
}