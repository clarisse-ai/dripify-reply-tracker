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
        "",
    "",
    JSON.stringify(data)
  ]);

    const lastRow = sheet.getLastRow();
  const statusCell = sheet.getRange(lastRow, 9);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["sent email", "replied", "applied", "onboarded"], true)
    .setAllowInvalid(false)
    .build();

  statusCell.setDataValidation(statusRule);
  statusCell.setValue("");


  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

 const STATUS_COLUMN = 9; // I
const CAMPAIGN_COLUMN = 3; // C
const LEAD_NAME_COLUMN = 4; // D
const EMAIL_COLUMN = 8; // H
const NOTES_COLUMN = 10; // J
const EMAIL_SENT_AT_COLUMN = 12; // L

function handleStatusEdit(e) {
  const sheet = e.range.getSheet();

  if (sheet.getName() !== "Sheet1") return;
  if (e.range.getColumn() !== STATUS_COLUMN) return;
  if (e.range.getRow() === 1) return;

  const newStatus = String(e.value || "").toLowerCase().trim();
  if (newStatus !== "sent email") return;

  const row = e.range.getRow();
  const emailSentAtCell = sheet.getRange(row, EMAIL_SENT_AT_COLUMN);

  if (emailSentAtCell.getValue()) return;

  const campaign = String(sheet.getRange(row, CAMPAIGN_COLUMN).getValue() || "");
  const leadName = String(sheet.getRange(row, LEAD_NAME_COLUMN).getValue() || "");
  const email = String(sheet.getRange(row, EMAIL_COLUMN).getValue() || "").trim();

  if (!email) {
    sheet.getRange(row, NOTES_COLUMN).setValue("Email not sent: missing email address.");
    return;
  }

  const template = getEmailTemplateForCampaign_(campaign, leadName);

  if (!template) {
    sheet.getRange(row, NOTES_COLUMN).setValue("Email not sent: no matching template for campaign.");
    return;
  }

  GmailApp.sendEmail(email, template.subject, template.body);

  emailSentAtCell.setValue(new Date());
}

function getEmailTemplateForCampaign_(campaign, leadName) {
  const campaignLower = campaign.toLowerCase();
  const firstName = getFirstName_(leadName);

  if (campaignLower.includes("accounting")) {
    return {
      subject: "Next steps for Fleet's Domain Expert position",
      body: `Hi ${firstName},

I’m excited to move forward with the next steps for our Domain Expert position.

At Fleet, we are hiring experts to help us train agents on the specific workflows that professionals in your field encounter every day. We want to leverage your deep industry knowledge to ensure that the applications we are testing are rooted in real-world accounting scenarios.

Regarding the time commitment, we are looking for at least 15–20 hours per week, though there is certainly the opportunity to work more if your schedule allows.

To get the formal process started, please submit your official application here: https://www.fleetai.com/careers/domain-expert-accounting?utm_source=dripify&utm_medium=outbound&utm_campaign=posting

Once that is in, I’ll reach out right away to schedule a time for us to talk. I’m looking forward to discussing your experience further and answering any questions you have about what we’re building at Fleet.

Best regards,
Clarisse, Operations`
    };
  }
if (campaignLower.includes("corpfin")) {
    return {
      subject: "Next steps for Fleet's Domain Expert position",
      body: `Hi ${firstName},

I’m excited to move forward with the next steps for our Domain Expert position.

At Fleet, we are hiring experts to help us train agents on the specific workflows that professionals in your field encounter every day. We want to leverage your deep industry knowledge to ensure that the applications we are testing are rooted in real-world corporate finance scenarios.

Regarding the time commitment, we are looking for at least 15-20 hours per week, though there is certainly the opportunity to work more if your schedule allows.

To get the formal process started, please submit your official application here: https://www.fleetai.com/careers/fleet-fellow-corp-finance?utm_source=dripify&utm_medium=outbound&utm_campaign=posting

Once that is in, I’ll reach out right away to schedule a time for us to talk. I’m looking forward to discussing your experience further and answering any questions you have about what we’re building at Fleet.

Best regards,
Clarisse, Operations`
    };
  }
    if (campaignLower.includes("insurance")) {
    return {
      subject: "Next steps for Fleet's Domain Expert position",
      body: `Hi ${firstName},

I’m excited to move forward with the next steps for our Domain Expert position.

At Fleet, we are hiring experts to help us train agents on the specific workflows that professionals in your field encounter every day. We want to leverage your deep industry knowledge to ensure that the applications we are testing are rooted in real-world insurance scenarios.

Regarding the time commitment, we are looking for at least 15–20 hours per week, though there is certainly the opportunity to work more if your schedule allows.

To get the formal process started, please submit your official application here: https://www.fleetai.com/careers/domain-expert-insurance?utm_source=dripify&utm_medium=outbound&utm_campaign=posting

Once that is in, I’ll reach out right away to schedule a time for us to talk. I’m looking forward to discussing your experience further and answering any questions you have about what we’re building at Fleet.


Best regards,
Clarisse, Operations`
    };
  }


  return null;
}

function getFirstName_(leadName) {
  const name = String(leadName || "").trim();
  if (!name) return "there";
  return name.split(" ")[0];
}
 

