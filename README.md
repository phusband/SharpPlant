![SharpPlant Logo](http://i1272.photobucket.com/albums/y394/allockse/Development/SharpPlant_zps70fa0e1c.jpeg)

SharpPlant is a .NET wrapper for Intergraph SmartPlant® products, written in C#.  It was designed to simplify .NET development, as well as provide additional API functionality as needed.  The initial release of this API covers SmartPlant® Review only.  However as time, resources and community involvement picks up, the roadmap is planned to expand into further Intergraph products.  You can view more about the Intergraph SmartPlant® product line at http://www.intergraph.com

## How to Get Started
- [Download](https://github.com/Allockse/SharpPlant/archive/master.zip) the latest SharpPlant project zip and build the dll in Visual Studio, or your preferred .NET IDE.
- Add the SharpPlant.dll as a reference to a new or existing .NET solution to make the classes/methods available.
- SharpPlant <b>requires .NET 4.0</b> or higher as it implements dynamic variables at runtime.

## Example Usage
### Creating an application reference:
``` csharp
 // Create a new application object
 var sprApp = new SprApplication();

```

### Updating the application text window:
``` csharp
 // Update the application text window
 sprApp.TextWindow_Update("SharpPlant loaded successfully");

```

### Placing a new tag:
``` csharp
 // Create a new tag object
 var newTag = new SprTag();

 // Set the tag properties
 newTag.Text = "High-point vent required as per P&ID";
 newTag.Discipline = "Piping";
 newTag.Creator = "ExampleUser";
            
 // Place the tag in the active application
 sprApp.Tags_Place(ref newTag);

```

### Placing a new annotation:
``` csharp
 // Create a new annotation object
 var newAnno = new SprAnnotation();

 // Set the annotation properties
 newAnno.Text = "Removable handrail";
 newAnno.DisplayBackground = true;
 newAnno.DisplayLeader = true;
 newAnno.BackgroundColor = System.Drawing.Color.Green;
 newAnno.Persistent = true;

 // Place the annotation in the active application
 sprApp.Annotations_Place(ref newAnno);

```

### Taking a snapshot:
``` csharp
 // Create a new snapshot object
 var newSnap = new SprSnapShot();

 // Set the snapshot properties
 newSnap.AntiAlias = 3;
 newSnap.OutputFormat = SprSnapshotFormat.Png;
 newSnap.Overwrite = true;

 // Take the snapshot
 sprApp.TakeSnapshot(newSnap, "TestImage", @"C:\TestSnapshots\");

```

### Collect object information:
``` csharp
 // Prompt a user to select an object in the active application instance
 var newObjectData = sprApp.GetObjectData("Select an object to view");

```

### DrAPI Object Wrapping:
- DrAPI: SprApplication
- DrAnnotationDbl: SprAnnotation
- DrDisplaySetDbl: NOT WRAPPED
- DrKey: SprLinkage
- DrLongArray: OBSOLETE
- DrMask: NOT WRAPPED
- DrMeasurement: NOT WRAPPED
- DrMeasurementCollection: NOT WRAPPED
- DrObjectDataDbl: SprObject
- DrPointDbl: SprPoint
- DrSnapShot: SprSnapShot
- DrStringArray: OBSOLETE
- DrTransform: NOT WRAPPED
- DrViewDbl: NOT WRAPPED
- DrVolumeAnnotation: NOT WRAPPED
- DrWindow: SprWindow

## SmartPlant® Review Development Notes:
SharpPlant does not require the VaxCtrl3.dll to be added as a project reference, however it needs to be registered
on any machine it runs on (installing SmartPlant® Review does this automatically).

Additional resources for SmartPlant® Review development can be found in the ..\resdlls\0009 directory inside the local install folder:
- DRAPI_API.cmh
- Drapix_API.chm

If the API module is installed, additional libraries and files will be located in the ..\Api directory inside the local install folder:
- DRAPI.H
- drapi32.dll
- drapi32.lib

## Release History
- 4/8/2013 Initial release
- 8/12/2013 Beta 0.4 Release

## Contributors
[Parrish Husband] (https://github.com/phusband)

## Support/Discussion
[Official support thread at DaveTyner.com] (http://www.davetyner.com/forum/forumdisplay.php?127-Sharp-Plant)

## License
SharpPlant is available under the MIT license. See the LICENSE file for more info.
