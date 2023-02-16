import { Context } from "@azure/functions";
import { createMicrosoftGraphClient, OnBehalfOfUserCredential, createMicrosoftGraphClientWithCredential, IdentityType, TeamsFx } from "@microsoft/teamsfx";

export async function getInstallationId(context: Context, teamsfx: TeamsFx, userId: string): Promise<any> {
  console.log("In getInstallationId");
  try {
    const teamsAppId = process.env.TEAMS_APP_ID;
    const apiPath =
      "/users/" +
      userId +
      "/teamwork/installedApps?$expand=teamsApp,teamsAppDefinition&$filter=teamsApp/externalId eq " +
      "'" + teamsAppId + "'";

    let teamsfx_app;
    teamsfx_app = new TeamsFx(IdentityType.App);
    const graphClient = createMicrosoftGraphClient(teamsfx_app, [".default"]);
    const appInstallationInfo = await graphClient.api(apiPath).get();
    const appArray = appInstallationInfo["value"][0];
    const installationId = appArray["id"];
    
    return installationId;
  } catch (e) {
    context.log.error(e);
    throw new Error("In getInstallationId" + e.message + ": " + e.response?.data?.error);
  }
}