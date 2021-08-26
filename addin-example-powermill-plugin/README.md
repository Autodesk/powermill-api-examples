# Example PowerMill Plugin
Example C# and VB projects for how to use PowerMill plugins.  This also includes an installer for the plugins.

# Setup
To set this up for development you will need:

- `PowerMill.dll`, this can be generated as described in the main repository's [README](/README.md).
- `PluginFramework.dll`, this is found in a PowerMill installation, as documented in the main repository's [README](/README.md).

Once you have these you should then be able to add a reference from your plugin to the PluginFramework solution and create your plugin from this base.

# Developing
When using this, please change the class Id of the plugin in the plugin class as well as in the RegisterDLL class otherwise everyone throughout the world will have the same Guid for their plugins and if you develop more than one plugin then they won't work together.
If you search the solution for "TODO" you will find all the places you need to make this change.
If you want to create an installer for the plugins, then remove the "Primary output from ExamplePowerMillPlugin (Active)" or "Primary output from ExamplePowerMillPluginVB (Active)" depending on which one you don't want to use.
