---
page_type: sample
languages:
- csharp
products:
- office-teams
description: Teams application to create assignments and distribute the team members in randomized groups
urlFragment: microsoft-teams-apps-groupactivities
---

| [Documentation](https://github.com/OfficeDev/microsoft-teams-apps-groupactivities/wiki) | [Deployment guide](https://github.com/OfficeDev/microsoft-teams-apps-groupactivities/wiki/Deployment-guide) | [Architecture](https://github.com/OfficeDev/microsoft-teams-apps-groupactivities/wiki/Solution-Overview)
|--|--|--|

## Group Activities App Template

For organizations including education institutions, there is often a need to create and collaborate in a group activity. Often, the process of creating and managing the groups and the collaboration around it takes more time than the activity itself. Group activities app makes it easy to quickly create groups for group activities along with providing an easy workflow to manage the collaboration all within the familiar context of Microsoft Teams canvas.
 
Without Group Activities app, it would involve first doing the pairing manually, then going and creating the channels (or letting team members figure their own collaboration environment) and finally adding users to such private channels.

Using the  app in Microsoft Teams, activity authors will have the ability to create assignments or activities and distribute the team members in a team/class in randomized groups depending on the grouping criteria and create Standard or Private channels for each of those groups that enables members to have their own workspace for collaboration. The activity author will be added to all standard channels created by the app.

The grouping criteria consists of two options
 - Create groups by dividing into equal number of team members in each group
 - Create groups by dividing team members into number of groups specified

While creating a group activity, activity authors will also have the option to select if they would like the app to send reminders until each activity is scheduled to be completed. The app will send 2 reminders daily at 10 AM and 5 PM until the due date of the activity.

![Group activity creating using messaging extension](https://github.com/OfficeDev/microsoft-teams-apps-groupactivities/wiki/Images/GroupActivities_01.png)

![Group activity create form](https://github.com/OfficeDev/microsoft-teams-apps-groupactivities/wiki/Images/GroupActivities_02.png)

![Activity summary message](https://github.com/OfficeDev/microsoft-teams-apps-groupactivities/wiki/Images/GroupActivities_03.png)


## Legal Notice
This app template is provided under the [MIT License](https://github.com/OfficeDev/microsoft-teams-apps-groupactivities/blob/master/LICENSE) terms.  In addition to these terms, by using this app template you agree to the following:

-	You are responsible for complying with all applicable privacy and security regulations related to use, collection and handling of any personal data by your app.  This includes complying with all internal privacy and security policies of your organization if your app is developed to be sideloaded internally within your organization.

-	Where applicable, you may be responsible for data related incidents or data subject requests for data collected through your app.

-	Any trademarks or registered trademarks of Microsoft in the United States and/or other countries and logos included in this repository are the property of Microsoft, and the license for this project does not grant you rights to use any Microsoft names, logos or trademarks outside of this repository.  Microsoft’s general trademark guidelines can be found [here](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general.aspx).

-	Use of this template does not guarantee acceptance of your app to the Teams app store.  To make this app available in the Teams app store, you will have to comply with the [submission and validation process](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/appsource/publish), and all associated requirements such as including your own privacy statement and terms of use for your app.

## Getting Started

Begin with the [Solution overview](https://github.com/OfficeDev/microsoft-teams-apps-groupactivities/wiki/Solution-Overview) to read about what the app does and how it works.

When you're ready to try out Group Activities app, or to use it in your own organization, follow the steps in the [Deployment guide](https://github.com/OfficeDev/microsoft-teams-apps-groupactivities/wiki/Deployment-guide).

## Feedback

Thoughts? Questions? Ideas? Share them with us on [Teams UserVoice](https://microsoftteams.uservoice.com/forums/555103-public) !

Please report bugs and other code issues [here](https://github.com/OfficeDev/microsoft-teams-apps-groupactivities/issues/new).

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
