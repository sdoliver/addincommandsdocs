Add-In Commands Preview Documentation
===================


Add-In Commands enable Office Web Add-Ins to **extend the Office user interface**. Developers define commands in the Add-In manifest and the platform creates native UX hooks depending on where the Add-In is running. 

For example,  an Outlook Add-In with commands in the MessageReadCommandSurface shows them as **Ribbon buttons** when running on the Win32 Office clients; when running on Outlook.com, commands will be interpreted and rendered in a way that fits Outlook.com. The result for end users is **deep integration into the Office user interface that provides efficient access to Add-In functionality**

This screenshot shows how Add-In commands look like when projected on the Outlook Desktop Ribbon. 
![](http://i.imgur.com/tqvb2Wl.png)

> **Note:**
> Currently commands are only supported on the Office Preview for Outlook Desktop(aka Win32). Support for other clients and platforms is in the works. 

How it works
-------------
Super simple. Your add-in defines its commands using the Add-In manifest. You then you publish you Add-in the same way you do today. Once the Add-In is added by a user or an administrator the Add-In platform interprets those commands and renders native UI hooks. When users trigger your command, the action specified in it is executed. 


This video shows you an E2E demo of Add-In commands
[![APP COMMANDS DEMO](http://i.imgur.com/T9jFQBi.png)](https://www.youtube.com/watch?v=jQGDZNsCkQs) 
Declaring Commands via the Add-In Manifest
-------------

The easiest way to get going is to use the **[sample manifest](https://github.com/LezaMax/addincommandsdocs/blob/master/manifestsamples/appCommandFullSampleMenu.xml)** as a template. From there you can tweak it to achieve the results you want. Below is a snippet that shows you how to add a button to the Outlook home that will appear when reading e-mails. 

     <ExtensionPoint xsi:type="MessageReadCommandSurface">	
        <!-- Tabs supported: <OfficeTab id="TabDefault">   OR <CustomTab id="[custom id]"> -->
          <OfficeTab id="TabDefault">

          	<!--If targetting an existing Office tab, only one group is supported. Custom tabs support multiple groups -->
            <Group id="msgreadTabMessage.grp1">
              <!-- Label must point to a short string resource -->
              <Label resid="groupLabel" />
               <!-- We have several bugs and redundancy on tooltips right now. Ignore them for the time being-->
              <Tooltip resid="residTipDescription" />

			<!--Control Types supported are "Button" and "Menu" but menu has bugs right now so stick with Buttons only-->
              <Control xsi:type="Button" id="button1id">
                <Label resid="residUILessButton" />
                <!-- We have several bugs and redundancy on tooltips right now. Ignore them for the time being-->
                <Tooltip resid="residTipDescription" />
                <Supertip>
                  <Title resid="residTipTitle" />
                  <Description resid="residTipDescription" />
                </Supertip>

                <!--Mandatory icons are 16,32 and 80. Other sizes supported are: . At runtime we pick the best (we have a bug that we aren't doing that right now)-->
                <Icon>
                  <bt:Image size="16" resid="functionIcon" />					
                  <bt:Image size="32" resid="functionIcon" />
                  <bt:Image size="80" resid="functionIcon" />
                </Icon>
                	<!--The ExecuteFunction action loads the FunctionFile and calls a function with the name provided -->
                    <Action xsi:type="ExecuteFunction">
                      <FunctionName>uiLessFunction</FunctionName>
                    </Action>
              </Control>
          </OfficeTab>
        </ExtensionPoint>

Below are the definitions for the most important nodes on the manifest

![](http://i.imgur.com/Z4L4dWT.png)

#### <i class="icon-file"></i> Hosts

Defines which Office application (Add-in host) is being extended. The only host that currently supports commands is **Mail** which maps to Outlook, Outlook.com and in the future Outlook for IOS and Android. 

#### <i class="icon-file"></i> FormFactor
This node groups all commands for a given form factor together so you can optimize your UI for different devices. Currently only the **Desktop** form factor is supported 

#### <i class="icon-file"></i> ExtensionPoint

Defines which part of the host the Add-In is extending. The following ExtensionPoints are supported:


 - MessageComposeCommandSurface
 - MessageReadCommandSurface
 - AppointmentOrganizerCommandSurface 
 - AppointmentAttendeeCommandSurface 


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

Testing Commands
----------
The easiest way to try Add-In commands is to upload them to your Exchange mailbox. Watch this video for an E2E demo: [https://www.youtube.com/watch?v=jQGDZNsCkQs](https://www.youtube.com/watch?v=jQGDZNsCkQs)



When And How to Use Add-In Commands
----------
We strongly encourage you to use commands as the main entry points to your Add-In functionality. 

####Few commands: Add them to existing tabs
If your main requirement is to provide quick access to launch your Add-In having a single command that launches your Add-In is probably your best alternative. If your require up to 6 commands then you can still add them to existing Outlook tabs by using the OfficeTab element as explained below. 
 
####Lots of commands: Create your own custom tab
If your Add-in provides enough functionality to require more than 6 commands, you might be better off creating your own custom tab. Use the CustomTab element as indicated in the sections below.  

Restrictions
----------
A single Add-in can only create:

- One custom tab AND/OR
- One group on an existing Office tab

A group created by an Add-in:

- Can only hold oup to 6 buttons
- Cannot specify its final layout or button size. Layout is automatically determined by the platform based on the button precedence and real state available. Buttons that are declared first on a group will generally be displayed in larger proportions that buttons declared at the end of a group.

At any point in time users or administrators can uninstall an Add-In and any UI that came with it will be removed.

  