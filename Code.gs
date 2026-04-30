/**
 * Google Apps Script webhook handler for Dripify first-reply tracking.
 *
 * Deploy this script as a Web App and use the Web App URL in Dripify campaign
 * webhooks. Add team/campaign labels as URL parameters, for example:
 *
 *   https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?teamMember=Sneha&campaign=Offer-Insurance
 */
function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  const receivedAt = new Date();

  let data = {};
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    data = { raw: e.postData.contents };
  }

  const teamMemberRaw =
    (e.parameter && e.parameter.teamMember) ||
    data.team_member ||
    data.sender ||
    data.account ||
    data.user ||
    data.member ||
    "";

  const teamMember = decodeURIComponent(teamMemberRaw);

  const campaignRaw =
    (e.parameter && e.parameter.campaign) ||
    data.campaign ||
    data.campaign_name ||
    data.campaignName ||
    data.sequence ||
    "";

  const campaign = decodeURIComponent(campaignRaw);

  const firstName = data.first_name || data.firstname || data.firstName || "";
  const lastName = data.last_name || data.lastname || data.lastName || "";
  const leadName =
    data.full_name ||
    data.name ||
    `${firstName} ${lastName}`.trim();

  const company = data.company || data.organization || "";
  const title = data.title || data.position || data.job_title || "";
  const linkedinUrl =
    data.linkedin_url ||
    data.profile_url ||
    data.linkedin ||
    data.link ||
    "";

  const email = data.email || "";

  sheet.appendRow([
    receivedAt,
    teamMember,
    campaign,
    leadName,
    company,
    title,
    linkedinUrl,
    email,
    "New",
    "",
    JSON.stringify(data)
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
