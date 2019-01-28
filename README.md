## DriveWorks Microsoft Dynamics 365 Integration Plugin

This source code has been made available to provide DriveWorks users with a plugin with which to build an integration between Microsoft Dynamics CRM and DriveWorks.

This code is provided under the MIT license,
for more details see [LICENSE.md](https://github.com/DriveWorks/Labs-Integration-Example/blob/master/LICENSE.md).

This plugin requires the DriveWorks 16 SDK or higher.

This plugin also has dependencies to the Microsoft CRM SDK. You should use the NuGet package manager to restore any references to this SDK.

### Building and Installation

1. Clone the repository into your local environment. 
2. Open in Visual Studio.
3. Restore any references to the Microsoft CRM SDK using the NuGet package manager.
4. Ensure any DriveWorks references are without error.
5. Build the entire solution.
6. If the solution built correctly it will output a file called 'DriveWorks.MSCRM.Integration.dll' to the bin directory.
7. Use this dll to install against DriveWorks http://docs.driveworkspro.com/Topic/HowToManuallyInstallAPlugin
8. Once the plugin is successfully installed. Configure the plugin settings to establish a connection to your CRM system. 

### Functionality

The plugin comes with 13 functions which pull data and 2 Specification Tasks which post data to Dynamics CRM.

Each Function and Specification Task comes with a full description available to view when building rules within DriveWorks, as well as commented public methods inside the source code. 

If you need any further clarity on the functionality of this code please contact apisupport@driveworks.co.uk

### Collaboration

This code base was developed for a specific internal integration at DriveWorks Ltd. The functionality was developed primarily for our needs in mind. The codebase barely scratches the surface of what is possible between Dynamics CRM and DriveWorks.

**Therefore users are invited to fork and submit pull requests to the code base.**

We also hope this plugin stands as a architectural blueprint of how to correctly design a DriveWorks plugin.
