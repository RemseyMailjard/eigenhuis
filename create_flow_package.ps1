# ─────────────────────────────────────────────────────────────────────────────
# create_flow_package.ps1  —  gebaseerd op een echte export als referentie
# ─────────────────────────────────────────────────────────────────────────────
param(
  [string]$Email = "remsey@skills4-it.nl"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$enc = [System.Text.UTF8Encoding]::new($false)

# ── Identifiers (zelfde patroon als demo: lowercase GUID in XML, UPPERCASE in bestandsnaam) ──
$flowGuidLower = "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
$flowGuidUpper = "A1B2C3D4-E5F6-7890-ABCD-EF1234567890"
$connRefLogical = "new_sharedoffice365_vehtr"
$flowJsonName = "VEHCopilotTraining-Aanmelding-$flowGuidUpper.json"

$workspace = "C:\Users\Remse\Desktop\Opdrachtgevers\E - Eigen huis"
$outZip = "$workspace\VEH_Copilot_Training_Flow.zip"
$tempDir = Join-Path $env:TEMP "VEH_SolPkg_$(Get-Random)"
$wfDir = "$tempDir\Workflows"

if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $wfDir | Out-Null

# ═══════════════════════════════════════════════════════════════════════════════
# [Content_Types].xml
# ═══════════════════════════════════════════════════════════════════════════════
$contentTypes = @'
<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml"  ContentType="application/octet-stream" />
  <Default Extension="json" ContentType="application/octet-stream" />
  <Default Extension="png"  ContentType="image/png" />
</Types>
'@

# ═══════════════════════════════════════════════════════════════════════════════
# solution.xml  —  alleen type 29 in RootComponents (geen connectionreference)
# ═══════════════════════════════════════════════════════════════════════════════
$solutionXml = @"
<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml version="9.2.24.10247" SolutionPackageVersion="9.2" languagecode="1033" generatedBy="CrmLive" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <SolutionManifest>
    <UniqueName>VEHCopilotTrainingAanmelding</UniqueName>
    <LocalizedNames>
      <LocalizedName description="VEH Copilot Training - Aanmelding" languagecode="1033" />
    </LocalizedNames>
    <Descriptions />
    <Version>1.0.0.0</Version>
    <Managed>0</Managed>
    <Publisher>
      <UniqueName>DefaultPublisherVEH</UniqueName>
      <LocalizedNames>
        <LocalizedName description="Default Publisher VEH" languagecode="1033" />
      </LocalizedNames>
      <Descriptions />
      <EMailAddress xsi:nil="true"></EMailAddress>
      <SupportingWebsiteUrl xsi:nil="true"></SupportingWebsiteUrl>
      <CustomizationPrefix>new</CustomizationPrefix>
      <CustomizationOptionValuePrefix>10000</CustomizationOptionValuePrefix>
      <Addresses>
        <Address>
          <AddressNumber>1</AddressNumber>
          <AddressTypeCode>1</AddressTypeCode>
          <City xsi:nil="true"></City>
          <County xsi:nil="true"></County>
          <Country xsi:nil="true"></Country>
          <Fax xsi:nil="true"></Fax>
          <FreightTermsCode xsi:nil="true"></FreightTermsCode>
          <ImportSequenceNumber xsi:nil="true"></ImportSequenceNumber>
          <Latitude xsi:nil="true"></Latitude>
          <Line1 xsi:nil="true"></Line1>
          <Line2 xsi:nil="true"></Line2>
          <Line3 xsi:nil="true"></Line3>
          <Longitude xsi:nil="true"></Longitude>
          <Name xsi:nil="true"></Name>
          <PostalCode xsi:nil="true"></PostalCode>
          <PostOfficeBox xsi:nil="true"></PostOfficeBox>
          <PrimaryContactName xsi:nil="true"></PrimaryContactName>
          <ShippingMethodCode>1</ShippingMethodCode>
          <StateOrProvince xsi:nil="true"></StateOrProvince>
          <Telephone1 xsi:nil="true"></Telephone1>
          <Telephone2 xsi:nil="true"></Telephone2>
          <Telephone3 xsi:nil="true"></Telephone3>
          <TimeZoneRuleVersionNumber xsi:nil="true"></TimeZoneRuleVersionNumber>
          <UPSZone xsi:nil="true"></UPSZone>
          <UTCOffset xsi:nil="true"></UTCOffset>
          <UTCConversionTimeZoneCode xsi:nil="true"></UTCConversionTimeZoneCode>
        </Address>
      </Addresses>
    </Publisher>
    <RootComponents>
      <RootComponent type="29" id="{$flowGuidLower}" behavior="0" />
    </RootComponents>
    <MissingDependencies />
  </SolutionManifest>
</ImportExportXml>
"@

# ═══════════════════════════════════════════════════════════════════════════════
# customizations.xml  —  structuur exact overgenomen van demo-export
# ═══════════════════════════════════════════════════════════════════════════════
$customizationsXml = @"
<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <Entities></Entities>
  <Roles></Roles>
  <Workflows>
    <Workflow WorkflowId="{$flowGuidLower}" Name="VEH Copilot Training - Aanmelding">
      <JsonFileName>/Workflows/$flowJsonName</JsonFileName>
      <Type>1</Type>
      <Subprocess>0</Subprocess>
      <Category>5</Category>
      <Mode>0</Mode>
      <Scope>4</Scope>
      <OnDemand>0</OnDemand>
      <TriggerOnCreate>0</TriggerOnCreate>
      <TriggerOnDelete>0</TriggerOnDelete>
      <AsyncAutodelete>0</AsyncAutodelete>
      <SyncWorkflowLogOnFailure>0</SyncWorkflowLogOnFailure>
      <StateCode>1</StateCode>
      <StatusCode>2</StatusCode>
      <RunAs>1</RunAs>
      <IsTransacted>1</IsTransacted>
      <IntroducedVersion>1.0</IntroducedVersion>
      <IsCustomizable>1</IsCustomizable>
      <BusinessProcessType>0</BusinessProcessType>
      <IsCustomProcessingStepAllowedForOtherPublishers>1</IsCustomProcessingStepAllowedForOtherPublishers>
      <ModernFlowType>0</ModernFlowType>
      <PrimaryEntity>none</PrimaryEntity>
      <LocalizedNames>
        <LocalizedName languagecode="1033" description="VEH Copilot Training - Aanmelding" />
      </LocalizedNames>
    </Workflow>
  </Workflows>
  <FieldSecurityProfiles></FieldSecurityProfiles>
  <Templates />
  <EntityMaps />
  <EntityRelationships />
  <OrganizationSettings />
  <optionsets />
  <CustomControls />
  <EntityDataProviders />
  <connectionreferences>
    <connectionreference connectionreferencelogicalname="$connRefLogical">
      <connectionreferencedisplayname>Office 365 Outlook</connectionreferencedisplayname>
      <connectorid>/providers/Microsoft.PowerApps/apis/shared_office365</connectorid>
      <iscustomizable>1</iscustomizable>
      <promptingbehavior>0</promptingbehavior>
      <statecode>0</statecode>
      <statuscode>1</statuscode>
    </connectionreference>
  </connectionreferences>
  <Languages>
    <Language>1033</Language>
  </Languages>
</ImportExportXml>
"@

# ═══════════════════════════════════════════════════════════════════════════════
# Workflow JSON  —  OpenApiConnection + parameters (zoals demo-export)
# Single-quoted here-string: PS expandeert geen $-variabelen binnenin.
# ═══════════════════════════════════════════════════════════════════════════════
$flowJson = @'
{
  "properties": {
    "connectionReferences": {
      "shared_office365": {
        "api": {
          "name": "shared_office365"
        },
        "connection": {
          "connectionReferenceLogicalName": "CONN_REF_LOGICAL"
        },
        "runtimeSource": "embedded"
      }
    },
    "definition": {
      "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
      "contentVersion": "1.0.0.0",
      "actions": {
        "Send_an_email_(V2)": {
          "type": "OpenApiConnection",
          "inputs": {
            "parameters": {
              "emailMessage/To": "TARGET_EMAIL",
              "emailMessage/Subject": "Nieuwe deelnemer Copilot-training: @{triggerBody()?['naam']}",
              "emailMessage/Body": "<table style=\"font-family:Calibri,Arial,sans-serif;font-size:14px;border-collapse:collapse;width:100%;max-width:600px\"><tr><td style=\"background:#3B2785;padding:20px 24px\"><span style=\"color:#F07800;font-size:20px;font-weight:bold\">&#127891;</span><span style=\"color:#ffffff;font-size:18px;font-weight:bold;margin-left:8px\">Nieuwe aanmelding \u2014 Copilot Training VEH</span></td></tr><tr><td style=\"padding:12px 24px;width:130px;font-weight:bold;color:#3B2785;border-bottom:1px solid #E5E7EB\">Naam</td><td style=\"padding:12px 24px;border-bottom:1px solid #E5E7EB\">@{triggerBody()?['naam']}</td></tr><tr style=\"background:#F9FAFB\"><td style=\"padding:12px 24px;font-weight:bold;color:#3B2785;border-bottom:1px solid #E5E7EB\">Functie</td><td style=\"padding:12px 24px;border-bottom:1px solid #E5E7EB\">@{triggerBody()?['functie']}</td></tr><tr><td style=\"padding:12px 24px;font-weight:bold;color:#3B2785;border-bottom:1px solid #E5E7EB\">E-mailadres</td><td style=\"padding:12px 24px;border-bottom:1px solid #E5E7EB\"><a href=\"mailto:@{triggerBody()?['email']}\" style=\"color:#3B2785\">@{triggerBody()?['email']}</a></td></tr><tr style=\"background:#F9FAFB\"><td style=\"padding:12px 24px;font-weight:bold;color:#3B2785;border-bottom:1px solid #E5E7EB;vertical-align:top\">Leerwens</td><td style=\"padding:12px 24px;border-bottom:1px solid #E5E7EB\">@{triggerBody()?['leerwens']}</td></tr><tr><td style=\"padding:12px 24px;font-weight:bold;color:#3B2785\">Datum &amp; tijd</td><td style=\"padding:12px 24px\">@{triggerBody()?['datum']} om @{triggerBody()?['tijdstip']}</td></tr><tr><td style=\"background:#FFF3E0;padding:12px 24px;font-size:11px;color:#9CA3AF;border-top:3px solid #F07800\">Verstuurd vanuit de Online Leeromgeving \u2022 Copilot in Dynamics 365 CE \u2022 Vereniging Eigen Huis</td></tr></table>",
              "emailMessage/Importance": "Normal"
            },
            "host": {
              "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
              "operationId": "SendEmailV2",
              "connectionName": "shared_office365"
            }
          },
          "runAfter": {}
        }
      },
      "triggers": {
        "manual": {
          "type": "Request",
          "kind": "Http",
          "inputs": {
            "triggerAuthenticationType": "All",
            "schema": {
              "type": "object",
              "properties": {
                "naam":     { "type": "string" },
                "functie":  { "type": "string" },
                "email":    { "type": "string" },
                "leerwens": { "type": "string" },
                "datum":    { "type": "string" },
                "tijdstip": { "type": "string" }
              }
            }
          }
        }
      },
      "parameters": {
        "$authentication": { "defaultValue": {}, "type": "SecureObject" },
        "$connections":    { "defaultValue": {}, "type": "Object" }
      },
      "outputs": {}
    },
    "templateName": null
  },
  "schemaVersion": "1.0.0.0"
}
'@
$flowJson = $flowJson.Replace("CONN_REF_LOGICAL", $connRefLogical).Replace("TARGET_EMAIL", $Email)

# ── Bestanden schrijven ───────────────────────────────────────────────────────
[System.IO.File]::WriteAllText("$tempDir\[Content_Types].xml", $contentTypes, $enc)
[System.IO.File]::WriteAllText("$tempDir\solution.xml", $solutionXml, $enc)
[System.IO.File]::WriteAllText("$tempDir\customizations.xml", $customizationsXml, $enc)
[System.IO.File]::WriteAllText("$wfDir\$flowJsonName", $flowJson, $enc)

# ── Zip maken ─────────────────────────────────────────────────────────────────
if (Test-Path $outZip) { Remove-Item $outZip -Force }
Compress-Archive -Path "$tempDir\*" -DestinationPath $outZip -CompressionLevel Optimal
Remove-Item $tempDir -Recurse -Force

Write-Host ""
Write-Host "Klaar!" -ForegroundColor Green
Write-Host "  Bestand   : $outZip" -ForegroundColor Cyan
Write-Host "  Ontvanger : $Email"  -ForegroundColor Cyan
Write-Host ""
Write-Host "Importeren: make.powerautomate.com > Solutions > Import Solution" -ForegroundColor Yellow
