# ─────────────────────────────────────────────────────────────────────────────
# create_flow_package.ps1
# Genereert een importeerbaar Power Automate SOLUTION-pakket (.zip)
# Vereiste structuur: [Content_Types].xml + solution.xml + customizations.xml
# Gebruik: .\create_flow_package.ps1 [-Email "jouw.naam@eigenhuis.nl"]
# ─────────────────────────────────────────────────────────────────────────────
param(
  [string]$Email = "remsey@skills4-it.nl"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$enc = [System.Text.UTF8Encoding]::new($false)   # UTF-8 zonder BOM

# ── Identifiers ───────────────────────────────────────────────────────────────
$flowGuidDashed = "a1b2c3d4-e5f6-7890-abcd-ef1234567890"   # lowercase: Dataverse vergelijkt case-sensitive
$flowGuidNoDash = "a1b2c3d4e5f67890abcdef1234567890"        # voor bestandsnaam
$flowSchemaName = "new_VEHCopilotTrainingAanmelding"
$connRefLogical = "new_sharedoffice365_vehtraining"
$workspace = "C:\Users\Remse\Desktop\Opdrachtgevers\E - Eigen huis"
$outZip = "$workspace\VEH_Copilot_Training_Flow.zip"
$tempDir = Join-Path $env:TEMP "VEH_SolPkg_$(Get-Random)"
$wfDir = "$tempDir\Workflows"

# ── Tijdelijke map ────────────────────────────────────────────────────────────
if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $wfDir | Out-Null

# ═══════════════════════════════════════════════════════════════════════════════
# 1. [Content_Types].xml
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
# 2. solution.xml
# ═══════════════════════════════════════════════════════════════════════════════
$solutionXml = @"
<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml version="9.2.24.10247" SolutionPackageVersion="9.2" languagecode="1033" generatedBy="CrmLive" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <SolutionManifest>
    <UniqueName>VEHCopilotTrainingAanmelding</UniqueName>
    <LocalizedNames>
      <LocalizedName description="VEH Copilot Training - Aanmelding" languagecode="1033" />
    </LocalizedNames>
    <Descriptions>
      <Description description="Ontvangt aanmeldingen van de online leeromgeving en stuurt een e-mailnotificatie naar de trainer." languagecode="1033" />
    </Descriptions>
    <Version>1.0.0.0</Version>
    <Managed>0</Managed>
    <Publisher>
      <UniqueName>DefaultPublisherVEH</UniqueName>
      <LocalizedNames>
        <LocalizedName description="Default Publisher VEH" languagecode="1033" />
      </LocalizedNames>
      <Descriptions />
      <EMailAddress xsi:nil="true" />
      <SupportingWebsiteUrl xsi:nil="true" />
      <CustomizationPrefix>new</CustomizationPrefix>
      <CustomizationOptionValuePrefix>10000</CustomizationOptionValuePrefix>
      <Addresses>
        <Address>
          <AddressNumber>1</AddressNumber>
          <AddressTypeCode>1</AddressTypeCode>
          <City xsi:nil="true" /><County xsi:nil="true" /><Country xsi:nil="true" />
          <Fax xsi:nil="true" /><FreightTermsCode xsi:nil="true" />
          <ImportSequenceNumber xsi:nil="true" /><Latitude xsi:nil="true" />
          <Line1 xsi:nil="true" /><Line2 xsi:nil="true" /><Line3 xsi:nil="true" />
          <Longitude xsi:nil="true" /><Name xsi:nil="true" /><PostalCode xsi:nil="true" />
          <PostOfficeBox xsi:nil="true" /><PrimaryContactName xsi:nil="true" />
          <ShippingMethodCode>1</ShippingMethodCode>
          <StateOrProvince xsi:nil="true" /><Telephone1 xsi:nil="true" />
          <Telephone2 xsi:nil="true" /><Telephone3 xsi:nil="true" />
          <TimeZoneRuleVersionNumber xsi:nil="true" /><UPSZone xsi:nil="true" />
          <UTCOffset xsi:nil="true" /><WebSiteURL xsi:nil="true" />
        </Address>
      </Addresses>
    </Publisher>
    <RootComponents>
      <RootComponent type="29"    id="{$flowGuidDashed}" behavior="0" />
      <RootComponent type="10466" schemaName="$connRefLogical" behavior="0" />
    </RootComponents>
    <MissingDependencies />
  </SolutionManifest>
</ImportExportXml>
"@

# ═══════════════════════════════════════════════════════════════════════════════
# 3. customizations.xml
# ═══════════════════════════════════════════════════════════════════════════════
$customizationsXml = @"
<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <Workflows>
    <Workflow WorkflowId="{$flowGuidDashed}" Name="VEH Copilot Training - Aanmelding" StateCode="1" StatusCode="2">
      <DesktopFlowModuleId xsi:nil="true" />
      <Category>5</Category>
      <Mode>0</Mode>
      <Scope>4</Scope>
      <OnDemand>1</OnDemand>
      <TriggerOnCreate>0</TriggerOnCreate>
      <TriggerOnDelete>0</TriggerOnDelete>
      <AsyncAutodelete>0</AsyncAutodelete>
      <SyncWorkflowLogOnFailure>0</SyncWorkflowLogOnFailure>
      <RunAs>1</RunAs>
      <IsTransacted>1</IsTransacted>
      <IntroducedVersion>1.0.0.0</IntroducedVersion>
      <IsCustomizable>1</IsCustomizable>
      <BusinessProcessType>0</BusinessProcessType>
      <IsCustomProcessingStepAllowedForOtherPublishers>1</IsCustomProcessingStepAllowedForOtherPublishers>
      <PrimaryEntity>none</PrimaryEntity>
      <LocalizedNames>
        <LocalizedName languagecode="1033" description="VEH Copilot Training - Aanmelding" />
      </LocalizedNames>
      <Descriptions>
        <Description languagecode="1033" description="Ontvangt aanmeldingen van de online leeromgeving en stuurt een e-mailnotificatie." />
      </Descriptions>
      <JsonFileName>/Workflows/${flowSchemaName}-${flowGuidNoDash}.json</JsonFileName>
    </Workflow>
  </Workflows>
  <connectionreferences>
    <connectionreference connectionreferencelogicalname="$connRefLogical">
      <connectionreferencedisplayname>Office 365 Outlook - VEH Training</connectionreferencedisplayname>
      <connectorid>/providers/Microsoft.PowerApps/apis/shared_office365</connectorid>
      <iscustomizable>1</iscustomizable>
      <statecode>0</statecode>
      <statuscode>1</statuscode>
    </connectionreference>
  </connectionreferences>
</ImportExportXml>
"@

# ═══════════════════════════════════════════════════════════════════════════════
# 4. Workflows/{schemaName}-{guidNoDash}.json  — volledige flow-definitie
#    Single-quoted here-string: PS expandeert GEEN $connections / $schema etc.
#    Daarna vervangen we de placeholders via .Replace().
# ═══════════════════════════════════════════════════════════════════════════════
$flowJson = @'
{
  "properties": {
    "connectionReferences": {
      "shared_office365": {
        "runtimeSource": "embedded",
        "connection": {
          "connectionReferenceLogicalName": "CONN_REF_LOGICAL"
        },
        "api": {
          "name": "shared_office365"
        }
      }
    },
    "definition": {
      "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
      "contentVersion": "1.0.0.0",
      "parameters": {
        "$connections":    { "defaultValue": {}, "type": "Object"       },
        "$authentication": { "defaultValue": {}, "type": "SecureObject" }
      },
      "triggers": {
        "manual": {
          "type": "Request",
          "kind": "Http",
          "inputs": {
            "method": "POST",
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
      "actions": {
        "Stuur_aanmeldingsmail": {
          "runAfter": {},
          "type": "ApiConnection",
          "inputs": {
            "host": {
              "connection": {
                "name": "@parameters('$connections')['shared_office365']['connectionId']"
              }
            },
            "method": "post",
            "path": "/v2/Mail",
            "body": {
              "To": "TARGET_EMAIL",
              "Subject": "Nieuwe deelnemer Copilot-training: @{triggerBody()?['naam']}",
              "Body": "<table style='font-family:Calibri,Arial,sans-serif;font-size:14px;border-collapse:collapse;width:100%;max-width:600px'><tr style='background:#3B2785'><td colspan='2' style='padding:16px 20px;color:white;font-size:18px;font-weight:bold'>&#127891; Nieuwe aanmelding &#8212; Copilot Training VEH</td></tr><tr><td style='padding:10px 20px;width:140px;font-weight:bold;border-bottom:1px solid #eee'>Naam</td><td style='padding:10px 20px;border-bottom:1px solid #eee'>@{triggerBody()?['naam']}</td></tr><tr style='background:#f9f9f9'><td style='padding:10px 20px;font-weight:bold;border-bottom:1px solid #eee'>Functie</td><td style='padding:10px 20px;border-bottom:1px solid #eee'>@{triggerBody()?['functie']}</td></tr><tr><td style='padding:10px 20px;font-weight:bold;border-bottom:1px solid #eee'>E-mailadres</td><td style='padding:10px 20px;border-bottom:1px solid #eee'>@{triggerBody()?['email']}</td></tr><tr style='background:#f9f9f9'><td style='padding:10px 20px;font-weight:bold;border-bottom:1px solid #eee'>Leerwens</td><td style='padding:10px 20px;border-bottom:1px solid #eee'>@{triggerBody()?['leerwens']}</td></tr><tr><td style='padding:10px 20px;font-weight:bold'>Datum &amp; tijd</td><td style='padding:10px 20px'>@{triggerBody()?['datum']} om @{triggerBody()?['tijdstip']}</td></tr><tr style='background:#FFF3E0'><td colspan='2' style='padding:10px 20px;font-size:12px;color:#888'>Verstuurd vanuit de Online Leeromgeving Copilot in Dynamics 365 CE &bull; Vereniging Eigen Huis</td></tr></table>",
              "Importance": "Normal",
              "IsHtml": true
            }
          }
        },
        "HTTP_respons": {
          "runAfter": {
            "Stuur_aanmeldingsmail": ["Succeeded", "Failed"]
          },
          "type": "Response",
          "kind": "Http",
          "inputs": {
            "statusCode": 200,
            "headers": {
              "Content-Type": "application/json",
              "Access-Control-Allow-Origin": "*"
            },
            "body": {
              "status": "ok",
              "message": "Aanmelding ontvangen"
            }
          }
        }
      },
      "outputs": {}
    }
  }
}
'@
$flowJson = $flowJson.Replace("CONN_REF_LOGICAL", $connRefLogical).Replace("TARGET_EMAIL", $Email)

# ── Bestanden schrijven ───────────────────────────────────────────────────────
[System.IO.File]::WriteAllText("$tempDir\[Content_Types].xml", $contentTypes, $enc)
[System.IO.File]::WriteAllText("$tempDir\solution.xml", $solutionXml, $enc)
[System.IO.File]::WriteAllText("$tempDir\customizations.xml", $customizationsXml, $enc)
[System.IO.File]::WriteAllText("$wfDir\${flowSchemaName}-${flowGuidNoDash}.json", $flowJson, $enc)

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
Write-Host "  1. Ga naar https://make.powerautomate.com" -ForegroundColor White
Write-Host "  2. Klik links op  Solutions  (Oplossingen)" -ForegroundColor White
Write-Host "  3. Klik op  Import Solution  >  Browse  > kies VEH_Copilot_Training_Flow.zip" -ForegroundColor White
Write-Host "  4. Klik op  Next  >  koppel je Office 365 Outlook-verbinding  >  Import" -ForegroundColor White
Write-Host "  5. Open de flow na import  >  kopieer de HTTP POST-URL uit de trigger" -ForegroundColor White
Write-Host "  6. Plak de URL in index.html bij:  const POWER_AUTOMATE_URL = ..." -ForegroundColor White
