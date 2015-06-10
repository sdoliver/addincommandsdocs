Add-In Commands Preview Documentation
===================


Add-In Commands enable Office Web Add-Ins to **extend the Office user interface**. Developers define commands in the Add-In manifest and the platform creates native UX hooks depending on where the Add-In is running. 

For example,  an Outlook Add-In with commands in the MessageReadCommandSurface shows them as **Ribbon buttons** when running on the Win32 Office clients; when running on Outlook.com, commands will be interpreted and rendered in a way that fits Outlook.com. The result for end users is **deep integration into the Office user interface that provides efficient access to Add-In functionality**
> **Note:**

> Currently commands are only supported on the Office Preview for Outlook Desktop(aka Win32). Support for other clients and platforms is in the works. 

How it works
-------------
Super simple. Your add-in defines its commands using the Add-In manifest. You then you publish you Add-in the same way you do today. Once the Add-In is added by a user or an administrator the Add-In platform interprets those commands and renders native UI hooks. When users trigger your command, the action specified in it is executed. 
[Image here]


Declaring Commands via the Add-In Manifest
-------------

The easiest way to get going is to use the **sample manifest** as a template. From there you can tweak it to achieve the results you want. Below is a snippet that shows you how to add a button to the Outlook home tab. 

    Manifest code goes here

Below are the definitions for the most important nodes on the manifest

[image here]![](http://i.imgur.com/Z4L4dWT.png)

#### <i class="icon-file"></i> Hosts

Defines which Office application (Add-in host) is being extended. The only host that currently supports commands is **Mail** which maps to Outlook, Outlook.com and in the future Outlook for IOS and Android. 

#### <i class="icon-file"></i> FormFactor
This node represents the form factor. Currently only the **Desktop** form factor is supported 

#### <i class="icon-file"></i> ExtensionPoint

Defines which part of the host the Add-In is extending. The following ExtensionPoints are supported:


 - MessageComposeCommandSurface
 - MessageReadCommandSurface
 - 

#### <i class="icon-file"></i> OfficeTab

Specifies an existing tab on the Office client to add commands to. The only value currently supported is "TabDefault" which adds commands to the default tab of the corresponding command surface. Adding commands to other existing tabs in Outlook is not supported.  

#### <i class="icon-file"></i> CustomTab

Specifies that your commands should be placed on a NEW tab. 

#### <i class="icon-file"></i> Group
Groups all your related commands together. 
....
#### <i class="icon-file"></i> Control
The actual UI affordance that will be used to display your command. Supported controls are:
- Button
- Menu

The sample manifest has examples for each. 
#### <i class="icon-file"></i> Action
The operation that will be excuted when users trigger your command (e.g. when they click on your button). Two main actions are supported:

- ShowTaskpaneUrl: Displays the given URL in a taskpane window 
- ExecuteFunction: Executes the given JavaScript function. The name that you provide must match the name of a functiona accesible in the global scope. The function is defined on a file referenced by the "FunctionFile" element on the manifest. 

####Resources node
All strings, urls and images used by commands are declared under the Resources node on the manifest. All other elements refer to these resources via the **resid** attribute. 

Testing and Debugging Commands
----------
[Todo: Add instructions here, or point to existing MSDN article on installing Outlook Add-Ins]

When And How to Use Add-In Commands
----------
We strongly encourage you to use commands as the main entry points to your Add-In functionality. 

####Few commands: Add them to existing tabs
If your main requirement is to provide quick access to launch your Add-In having a single command that launches your Add-In is probably your best alternative. If your require up to 6 commands then you can still add them to existing Outlook tabs by using the OfficeTab element as explained below. 
 
####Lots of commands: Create your own custom tab
If your Add-in provides enough functionality to require more than 6 commands, you might be better off creating your own custom tab. Use the CustomTab element as indicated in the sections below.  

Restrictions
----------
A single add-in can only create:

- One custom tab AND/OR
- One group on an existing Office tab

A group created by an Add-in:

- Can only hold oup to 6 buttons
- Cannot specify its final layout or button size. Layout is automatically determined by the platform based on the button precedence and real state available. Buttons that are declared first on a group will generally be displayed in larger proportions that buttons declared at the end of a group.

At any point in time users or administrators can uninstall an Add-In and any UI that came with it will be removed.

UI Projections Per Platform
----------
[Todo: Add table showing how the same command surface is displayed across platforms]
  