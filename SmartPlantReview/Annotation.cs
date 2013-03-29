using System;
using System.Drawing;

namespace SharpPlant.SmartPlantReview
{
    /// <summary>
    ///     Provides the properties for controlling text annotations in SmartPlant Review.
    /// </summary>
    public class Annotation
    {
        #region Annotation Properties

        /// <summary>
        ///     Active COM reference to the DrAnnotationDbl class.
        /// </summary>
        internal dynamic DrAnnotationDbl;

        /// <summary>
        ///     The parent Application reference.
        /// </summary>
        public Application Application { get; private set; }

        /// <summary>
        ///     Annotation background color if visible (0BGR format).
        /// </summary>
        public Color BackgroundColor
        {
            get { return IsActive ? SmartPlantReview.From0Bgr(DrAnnotationDbl.BackgroundColor) : Color.Empty; }
            set { if (IsActive) DrAnnotationDbl.BackgroundColor = SmartPlantReview.Get0Bgr(value); }
        }

        /// <summary>
        ///     Leader line color.
        /// </summary>
        public Color LineColor
        {
            get { return IsActive ? SmartPlantReview.From0Bgr(DrAnnotationDbl.LineColor) : Color.Empty; }
            set { if (IsActive) DrAnnotationDbl.LineColor = SmartPlantReview.Get0Bgr(value); }
        }

        /// <summary>
        ///     Annotation text color.
        /// </summary>
        public Color TextColor
        {
            get { return IsActive ? SmartPlantReview.From0Bgr(DrAnnotationDbl.TextColor) : Color.Empty; }
            set { if (IsActive) DrAnnotationDbl.TextColor = SmartPlantReview.Get0Bgr(value); }
        }

        /// <summary>
        ///     Center point of the annotation object.
        /// </summary>
        public Point3D CenterPoint
        {
            get
            {
                if (IsActive)
                    return new Point3D(DrAnnotationDbl.CenterPoint.East,
                                       DrAnnotationDbl.CenterPoint.North,
                                       DrAnnotationDbl.CenterPoint.Elevation);
                return null;
            }
            set
            {
                if (!IsActive) return;
                DrAnnotationDbl.CenterPoint.East = value.East;
                DrAnnotationDbl.CenterPoint.North = value.North;
                DrAnnotationDbl.CenterPoint.Elevation = value.Elevation;
            }
        }

        /// <summary>
        ///     End point of the leader line.
        /// </summary>
        public Point3D LeaderPoint
        {
            get
            {
                if (IsActive)
                    return new Point3D(DrAnnotationDbl.LeaderPoint.East,
                                       DrAnnotationDbl.LeaderPoint.North,
                                       DrAnnotationDbl.LeaderPoint.Elevation);
                return null;
            }
            set
            {
                if (!IsActive) return;
                DrAnnotationDbl.LeaderPoint.East = value.East;
                DrAnnotationDbl.LeaderPoint.North = value.North;
                DrAnnotationDbl.LeaderPoint.Elevation = value.Elevation;
            }
        }

        /// <summary>
        ///     Indicates whether a leader will be present.
        /// </summary>
        public bool DisplayLeader
        {
            get
            {
                // Return the bitwise zero check
                return (DrAnnotationDbl.Flags & Constants.ANNO_LEADER) != 0;
            }
            set
            {
                // Set flag true/false
                if (value) DrAnnotationDbl.Flags |= Constants.ANNO_LEADER;
                else DrAnnotationDbl.Flags &= ~Constants.ANNO_LEADER;
            }
        }

        /// <summary>
        ///     Indicates whether an arrowhead will be present.
        /// </summary>
        public bool DisplayArrowhead
        {
            get
            {
                // Return the bitwise zero check
                return (DrAnnotationDbl.Flags & Constants.ANNO_ARROW) != 0;
            }
            set
            {
                // Set flag true/false
                if (value) DrAnnotationDbl.Flags |= Constants.ANNO_ARROW;
                else DrAnnotationDbl.Flags &= ~Constants.ANNO_ARROW;
            }
        }

        /// <summary>
        ///     Indicates whether the background will be displayed.
        /// </summary>
        public bool DisplayBackground
        {
            get
            {
                // Return the bitwise zero check
                return (DrAnnotationDbl.Flags & Constants.ANNO_BACKGROUND) != 0;
            }
            set
            {
                // Set flag true/false
                if (value) DrAnnotationDbl.Flags |= Constants.ANNO_BACKGROUND;
                else DrAnnotationDbl.Flags &= ~Constants.ANNO_BACKGROUND;
            }
        }

        /// <summary>
        ///     Indicates that the annotation should be maintained between sessions.
        /// </summary>
        public bool SaveToMdb
        {
            get
            {
                // Return the bitwise zero check
                return (DrAnnotationDbl.Flags & Constants.ANNO_PERSIST) != 0;
            }
            set
            {
                // Set flag true/false
                if (value) DrAnnotationDbl.Flags |= Constants.ANNO_PERSIST;
                else DrAnnotationDbl.Flags &= ~Constants.ANNO_PERSIST;
            }
        }

        /// <summary>
        ///     Annotation text.
        /// </summary>
        public string Text
        {
            get { return IsActive ? DrAnnotationDbl.Text : null; }
            set { if (IsActive) DrAnnotationDbl.Text = value; }
        }

        /// <summary>
        ///     The session unique ObjectID of the annotation.
        /// </summary>
        public int AnnotationId { get; internal set; }

        /// <summary>
        ///     The object associated with the annotation.
        /// </summary>
        public ObjectData AssociatedObject { get; internal set; }

        /// <summary>
        ///     Annotation type.
        /// </summary>
        public string Type { get; set; }

        // Determines if a reference to the COM object is established
        private bool IsActive
        {
            get { return (DrAnnotationDbl != null); }
        }

        #endregion

        // Annotation initializer
        public Annotation()
        {
            // Link the parent application
            Application = SmartPlantReview.ActiveApplication;

            // Get a new DrAnnotationDbl object
            DrAnnotationDbl = Activator.CreateInstance(ImportedTypes.DrAnnotationDbl);

            // Set the default flags
            DrAnnotationDbl.Flags = 5;

            // Set the default annotation type
            Type = "Standard";

            // Set the default colors
            BackgroundColor = SmartPlantReview.From0Bgr(8454143);
            LineColor = SmartPlantReview.From0Bgr(8454143);
            TextColor = SmartPlantReview.From0Bgr(0);
        }
    }
}