# PowerMill to Vericut

Addin for PowerMill that exports stock, clamps, workpiece, tools, nc code and other required data from PowerMill to VERICUT.

## Setup

To set this up for development you will need:

- `MR_wrapper.dll`, this can be requested from CGTech (Vericut). This should be placed in the `PowerMILLVERICUTInterfacePlugin` folder.
- `PluginFramework.dll`, this is found in a PowerMill installation, as documented in the main repository's [README](/README.md). It should be copied to the `PowerMILLVERICUTInterfacePlugin` and `PowerMILLExporter`  folders.

You can now open the Visual Studio solution and build it.

## Developing

Once you have built the PowerMILLVERICUTInterfacePlugin solution, follow the instructions for **Running a Plugin** in the main repository's [README](/README.md). It is not necessary to follow the instructions in **Making the Plugin Accessible to PowerMill** if you have not changed the plugin class GUID, and have a version of the VERICUT Interface plugin installed (if you uninstall it you will need to follow those instructions).
