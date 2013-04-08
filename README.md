# SharpPlant
SharpPlant is a .NET wrapper for Intergraph SmartPlant® products, written in C#.  It was designed to simplify .NET development, as well as provide additional API functionality as needed.  The initial release of this API covers SmartPlant® Review only.  However as time, resources and community involvement picks up, the roadmap is planned to expand into further Intergraph products.  You can view more about the Intergraph SmartPlant® product line at http://www.intergraph.com

## How to Get Started
- [Download](https://github.com/Allockse/SharpPlant/archive/master.zip) the latest SharpPlant project zip and build the dll in Visual Studio.
- Add the ShartPlant.dll as a reference to a new or existing .NET solution to make the classes/methods available.

## Example Usage
### Creating an application reference:
``` csharp
// Create a new application object
var sprApp = new SprApplication();

```

### Updating the application text window:
``` csharp
// Update the application text window
sprAPp.TextWindow_Update("SharpPlant loaded successfully");

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
var newObjectData = sprApp.GetObjectData("Select an object to collect label information");

```

## Release History
- 4/8/2013 Initial release

## Contributors
[Parrish Husband] (https://github.com/Allockse)

## Support/Discussion
[Official thread on Dave Tyner] (http://www.davetyner.com/forum/)

## License
SharpPlant is available under the MIT license. See the LICENSE file for more info.

