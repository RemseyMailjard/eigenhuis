# ─────────────────────────────────────────────────────────────────────────────
# create_evaluatie_flow.ps1
# Genereert een importeerbaar Power Automate SOLUTION-pakket voor evaluaties.
# Gebruik: .\create_evaluatie_flow.ps1 [-Email "jouw@email.nl"]
# ─────────────────────────────────────────────────────────────────────────────
param(
    [string]$Email = "remsey@skills4-it.nl"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$enc = [System.Text.UTF8Encoding]::new($false)

# ── Identifiers (lowercase in XML, UPPERCASE in bestandsnaam — zie SKILL.md) ─
$flowGuidLower = "c3d4e5f6-a7b8-9012-cdef-123456789012"
$flowGuidUpper = "C3D4E5F6-A7B8-9012-CDEF-123456789012"
$connRefLogical = "new_sharedoffice365_evalu2"
$flowJsonName = "VEHCopilotTraining-Evaluatie-$flowGuidUpper.json"

$workspace = "C:\Users\Remse\Desktop\Opdrachtgevers\E - Eigen huis"
$outZip = "$workspace\VEH_Copilot_Evaluatie_Flow.zip"
$tempDir = Join-Path $env:TEMP "VEH_EvalPkg_$(Get-Random)"
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
# solution.xml — alleen type 29 in RootComponents (geen connectionreference)
# ═══════════════════════════════════════════════════════════════════════════════
$solutionXml = @"
<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml version="9.2.24.10247" SolutionPackageVersion="9.2" languagecode="1033" generatedBy="CrmLive" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <SolutionManifest>
    <UniqueName>VEHCopilotTrainingEvaluatie</UniqueName>
    <LocalizedNames>
      <LocalizedName description="VEH Copilot Training - Evaluatie" languagecode="1033" />
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
# customizations.xml — exacte structuur van een echte Dataverse-export
# ═══════════════════════════════════════════════════════════════════════════════
$customizationsXml = @"
<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <Entities></Entities>
  <Roles></Roles>
  <Workflows>
    <Workflow WorkflowId="{$flowGuidLower}" Name="VEH Copilot Training - Evaluatie">
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
        <LocalizedName languagecode="1033" description="VEH Copilot Training - Evaluatie" />
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
      <connectionreferencedisplayname>Office 365 Outlook - VEH Evaluatie</connectionreferencedisplayname>
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
# Workflow JSON — OpenApiConnection, exacte structuur van demo-export
# Single-quoted here-string: PS expandeert geen $-variabelen hierin.
# CONN_REF_LOGICAL en TARGET_EMAIL worden vervangen via .Replace()
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
        "Send_evaluatie_email": {
          "type": "OpenApiConnection",
          "inputs": {
            "parameters": {
              "emailMessage/To": "TARGET_EMAIL",
              "emailMessage/Subject": "Evaluatie Copilot Training: @{triggerBody()?['naam']} \u2014 @{triggerBody()?['beoordeling']}/5 sterren",
              "emailMessage/Body": "<table style='font-family:Calibri,Arial,sans-serif;font-size:14px;border-collapse:collapse;width:100%;max-width:640px'><tr><td colspan='2' style='background:#3B2785;padding:20px 24px'><span style='color:#F07800;font-size:22px;font-weight:800'>&#9733;</span><span style='color:#ffffff;font-size:18px;font-weight:700;margin-left:8px'>Trainings-evaluatie \u2014 Copilot Training VEH</span></td></tr><tr style='background:#FFF8F0'><td colspan='2' style='padding:14px 24px;border-bottom:2px solid #F07800'><span style='font-size:28px;font-weight:800;color:#3B2785'>@{triggerBody()?['beoordeling']}</span><span style='font-size:16px;color:#777;margin-left:6px'>/ 5 sterren</span></td></tr><tr><td style='padding:11px 24px;width:170px;font-weight:700;color:#3B2785;border-bottom:1px solid #E5E7EB'>Naam</td><td style='padding:11px 24px;border-bottom:1px solid #E5E7EB'>@{triggerBody()?['naam']}</td></tr><tr style='background:#F9FAFB'><td style='padding:11px 24px;font-weight:700;color:#3B2785;border-bottom:1px solid #E5E7EB'>Functie</td><td style='padding:11px 24px;border-bottom:1px solid #E5E7EB'>@{triggerBody()?['functie']}</td></tr><tr><td style='padding:11px 24px;font-weight:700;color:#3B2785;border-bottom:1px solid #E5E7EB;vertical-align:top'>Meest waardevol</td><td style='padding:11px 24px;border-bottom:1px solid #E5E7EB'>@{triggerBody()?['meest_waardevol']}</td></tr><tr style='background:#F9FAFB'><td style='padding:11px 24px;font-weight:700;color:#3B2785;border-bottom:1px solid #E5E7EB;vertical-align:top'>Verbeterpunten</td><td style='padding:11px 24px;border-bottom:1px solid #E5E7EB'>@{if(empty(triggerBody()?['verbeterpunten']), '(geen)', triggerBody()?['verbeterpunten'])}</td></tr><tr><td style='padding:11px 24px;font-weight:700;color:#3B2785;border-bottom:1px solid #E5E7EB'>Aanbevelen</td><td style='padding:11px 24px;border-bottom:1px solid #E5E7EB'>@{triggerBody()?['aanbevelen']}</td></tr><tr style='background:#F9FAFB'><td style='padding:11px 24px;font-weight:700;color:#3B2785;border-bottom:1px solid #E5E7EB;vertical-align:top'>Opmerkingen</td><td style='padding:11px 24px;border-bottom:1px solid #E5E7EB'>@{if(empty(triggerBody()?['opmerkingen']), '(geen)', triggerBody()?['opmerkingen'])}</td></tr><tr><td style='padding:11px 24px;font-weight:700;color:#3B2785'>Datum &amp; tijd</td><td style='padding:11px 24px'>@{triggerBody()?['datum']} om @{triggerBody()?['tijdstip']}</td></tr><tr><td colspan='2' style='background:#FFF3E0;padding:10px 24px;font-size:11px;color:#9CA3AF;border-top:3px solid #F07800'>Verstuurd vanuit de Evaluatiepagina \u2022 Copilot Training \u2022 Vereniging Eigen Huis</td></tr></table>",
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
                "naam":            { "type": "string" },
                "functie":         { "type": "string" },
                "beoordeling":     { "type": "string" },
                "meest_waardevol": { "type": "string" },
                "verbeterpunten":  { "type": "string" },
                "aanbevelen":      { "type": "string" },
                "opmerkingen":     { "type": "string" },
                "datum":           { "type": "string" },
                "tijdstip":        { "type": "string" }
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

# ── Schrijf bestanden ─────────────────────────────────────────────────────────
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
Write-Host "Importeren in Power Automate:" -ForegroundColor Yellow
Write-Host "  1. make.powerautomate.com > Solutions > Import Solution" -ForegroundColor White
Write-Host "  2. Kies VEH_Copilot_Evaluatie_Flow.zip > Next" -ForegroundColor White
Write-Host "  3. Koppel Office 365 Outlook-verbinding > Import" -ForegroundColor White
Write-Host "  4. Open de flow > kopieer de HTTP POST-URL uit de trigger" -ForegroundColor White
Write-Host "  5. Plak de URL in evaluatie.html bij: const POWER_AUTOMATE_URL = ..." -ForegroundColor White
