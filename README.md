# PowerMill API Examples
Welcome to the Autodesk PowerMill API Examples repository.   This repo provides the source code for advanced plugins which are included with the PowerMill application only in compiled form or not included directly at all.  These plugins are translators that use the PowerMill API to export data from PowerMill to third party applications such as Vericut.

The PowerMill development team is not actively maintaining or enhancing these plugins.  The source code is being made available to third parties who would like to do their own enhancements, or who would like to use these examples as templates for creating interfaces to other products.  Contributions to this repo will be reviewed and considered for inclusion (see below), but please be aware that these plugins have low development priority.

To learn more about PowerMill, visit the forum [here](https://forums.autodesk.com/t5/powermill-forum/bd-p/280) or the product website [here](https://www.autodesk.com/products/powermill/overview).

The repository for the PowerMill API Examples is hosted at:
https://github.com/Autodesk/powermill-api-examples

To Contribute please see [Contribute.md](Contribute.md). 

The extension is distributed under the Apache 2.0 license. See [LICENSE.txt](LICENSE.txt).

## Building the Plugins

There are just a few simple steps for building one of these plugins:

1. Clone the repository.
2. For the plugin you want to build, open the solution in Visual Studio.
3. Obtain dependencies (you will need `PowerMill.dll`, `PluginFramework.dll` and possibly others detailed in the specific README).
4. Build.

The plugins may have more specific instructions in their README. 

### Generating `PowerMill.dll`

The plugins will require that you generate `PowerMill.dll`, which contains the interface that they need to communicate with PowerMill. The specific README will contain instructions of what to do with this.

To generate `PowerMill.dll` follow these instructions:

- Using a standard command window, move to the PowerMILL's executable directory. e.g.
```
cd "C:\Program Files\Autodesk\PowerMill 2019\sys\exec64"
```
- Add the Microsoft SDK tools to the PATH variable if it's not already present (the exact path will different if you're not using .NET 4.0):
```
SET PATH=%PATH%;C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\bin\
```
- Generate a strong name key file using the strong name tool sn.exe:
```
sn.exe -k key_pair.snk
```
- Create the DLL from the type library using the `/keyfile` and `/asmversion` options (use the correct version number for the version of PowerMILL that's installed) e.g.
```
tlbimp.exe pmill.exe /keyfile:key_pair.snk /asmversion:2019.1.6
```
- The output file is called `PowerMILL.dll` as no output file was specified.

### Finding `PluginFramework.dll`

`PluginFramework.dll` is included in each install of PowerMill. It is found at 
```
<Install-Location>\file\plugins\framework\PluginFramework40\bin\Release
```
e.g. 
```
C:\Program Files\Autodesk\PowerMill 2019\file\plugins\framework\PluginFramework40\bin\Release
```

## Running the Plugins 
To run the plugin within PowerMill, you must COM register it using `regasm.exe`.

### Registering the COM component

To register the plugin, from the plugins's 'bin' directory:

```
C:\WINDOWS\Microsoft.NET\Framework64\v4.0.30319\regasm.exe BasicPlugin.dll /register /codebase
```

Note: This will issue a warning about the `/codebase` option being used not in conjunction with a strong named assembly. It is safe to ignore this warning, although giving your assembly a strong name will prevent the warning.

### Running the Plugin

After registering your plugin, you will be able to access it from within PowerMill by following these steps.

- Run PowerMill
- Press the "File" button on the top left of the ribbon.
- Open the Plugin Manager form, by going to "Options/Manage Installed Plugins"
- Select your plugin in the list, and select Enable.
- Your plugin should now have a green status icon, and be enabled.

## Contributions
In order to clarify the intellectual property license granted with Contributions from any person or entity, Autodesk must have a Contributor License Agreement ("CLA") on file that has been signed by each Contributor to this Open Source Project (the “Project”), indicating agreement to the license terms. This license is for your protection as a Contributor to the Project as well as the protection of Autodesk and the other Project users; it does not change your rights to use your own Contributions for any other purpose. There is no need to fill out the agreement until you actually have a contribution ready. Once you have a contribution you simply fill out and sign the applicable agreement (see the contributor folder in the repository) and send it to us at the address in the agreement.

## Trademarks

The license does not grant permission to use the trade names, trademarks, service marks, or product names of Autodesk, except as required for reasonable and customary use in describing the origin of the work and reproducing the content of any notice file. Autodesk, the Autodesk logo, Inventor HSM, HSMWorks, HSMXpress, Fusion 360, PowerMill, PartMarker, and PowerMILL are registered trademarks or trademarks of Autodesk, Inc., and/or its subsidiaries and/or affiliates in the USA and/or other countries. All other brand names, product names, or trademarks belong to their respective holders. Autodesk is not responsible for typographical or graphical errors that may appear in this document.
