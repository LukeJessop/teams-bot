const notificationTemplate = require("./adaptiveCards/notification-default.json");
const { notificationApp } = require("./internal/initialize");
const ACData = require("adaptivecards-templating");
const { teamsBot } = require("./teamsBot");
const express = require("express");

// Create express application.
const expressApp = express();
expressApp.use(express.json());

const server = expressApp.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${expressApp.name} listening to`, server.address());
});

// HTTP trigger to send notification. You need to add authentication / authorization for this API. Refer https://aka.ms/teamsfx-notification for more details.
expressApp.post("/api/notification", async (req, res) => {
  const pageSize = 100;
  let continuationToken = undefined;
  do {
    const pagedData = await notificationApp.notification.getPagedInstallations(
      pageSize,
      continuationToken
    );
    const installations = pagedData.data;
    continuationToken = pagedData.continuationToken;

    for (const target of installations) {
      await target.sendAdaptiveCard(
        new ACData.Template(notificationTemplate).expand({
          $root: {
            title: "New Event Occurred!",
            appName: "Contoso App Notification",
            description: `This is a sample http-triggered notification to ${target.type}`,
            notificationUrl: "https://aka.ms/teamsfx-notification-new",
          },
        })
      );

      /****** To distinguish different target types ******/
      /** "Channel" means this bot is installed to a Team (default to notify General channel)
        if (target.type === NotificationTargetType.Channel) {
          // Directly notify the Team (to the default General channel)
          await target.sendAdaptiveCard(...);

          // List all channels in the Team then notify each channel
          const channels = await target.channels();
          for (const channel of channels) {
            await channel.sendAdaptiveCard(...);
          }

          // List all members in the Team then notify each member
          const pageSize = 100;
          let continuationToken = undefined;
          do {
            const pagedData = await target.getPagedMembers(pageSize, continuationToken);
            const members = pagedData.data;
            continuationToken = pagedData.continuationToken;

            for (const member of members) {
              await member.sendAdaptiveCard(...);
            }
          } while (continuationToken);
        }
        **/

      /** "Group" means this bot is installed to a Group Chat
        if (target.type === NotificationTargetType.Group) {
          // Directly notify the Group Chat
          await target.sendAdaptiveCard(...);

          // List all members in the Group Chat then notify each member
          const pageSize = 100;
          let continuationToken = undefined;
          do {
            const pagedData = await target.getPagedMembers(pageSize, continuationToken);
            const members = pagedData.data;
            continuationToken = pagedData.continuationToken;

            for (const member of members) {
              await member.sendAdaptiveCard(...);
            }
          } while (continuationToken);
        }
        **/

      /** "Person" means this bot is installed as a Personal app
        if (target.type === NotificationTargetType.Person) {
          // Directly notify the individual person
          await target.sendAdaptiveCard(...);
        }
        **/
    }
  } while (continuationToken);

  /** You can also find someone and notify the individual person
    const member = await notificationApp.notification.findMember(
      async (m) => m.account.email === "someone@contoso.com"
    );
    await member?.sendAdaptiveCard(...);
    **/

  /** Or find multiple people and notify them
    const members = await notificationApp.notification.findAllMembers(
      async (m) => m.account.email?.startsWith("test")
    );
    for (const member of members) {
      await member.sendAdaptiveCard(...);
    }
    **/

  res.json({});
});

// Microsoft 365 Agents SDK message handler.
expressApp.post("/api/messages", async (req, res) => {
  await notificationApp.requestHandler(req, res, async (context) => {
    await teamsBot.run(context);
  });
});
