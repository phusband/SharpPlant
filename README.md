# SharpPlant
SharpPlant is a .NET wrapper for Intergraph SmartPlant® products, written in C#.  It was designed to simplify .NET development, as well as provide additional API functionality as needed.  The initial release of this API covers SmartPlant® Review only.  However as time, resources and community involvement picks up, the roadmap is planned to expand into further Intergraph products.  You can view more about the Intergraph SmartPlant® product line at http://www.intergraph.com

## How to Get Started
- [Download](https://github.com/Allockse/SharpPlant/archive/master) the latest SharpPlant project zip and build the dll in Visual Studio.
- Add the ShartPlant.dll as a reference to a new or existing .NET solution to make the classes/methods available.

## Example Usage
``` c-sharp
var sprApp = new SprApplication();
if (sprApp.IsConnected)
{
  sprApp.TextWindow_Update("SharpPlant loaded!");
}

```
