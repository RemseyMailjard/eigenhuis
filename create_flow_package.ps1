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
$flowGuidDashed  = "A1B2C3D4-E5F6-7890-ABCD-EF1234567890"   # voor XML WorkflowId
$flowGuidNoDash  = "A1B2C3D4E5F67890ABCDEF1234567890"        # voor bestandsnaam
$flowSchemaName  = "new_VEHCopilotTrainingAanmelding"
$connRefLogical  = "new_sharedoffice365_vehtraining"
$workspace       = "C:\Users\Remse\Desktop\Opdrachtgevers\E - Eigen huis"
$outZip          = "$workspace\VEH_Copilot_Training_Flow.zip"
$tempDir         = Join-Path $env:TEMP "VEH_SolPkg_$(Get-Random)"
$wfDir           = "$tempDir\Workflows"

# ── Tijdelijke map ────────────────────────────────────────────────────────────
if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $wfDir | Out-Null

# ── manifest.json (root van het pakket) ───────────────────────────────────────
# Beschrijft de inhoud van het pakket en declareert de verbindingen die nodig zijn.
$manifest = @'
{
  "schema": "1.0",
  "details": {
    "displayName": "VEH Copilot Training \u2014 Aanmelding",
    "description": "Ontvangt aanmeldingen van de online leeromgeving en stuurt een e-mailnotificatie naar de trainer.",
    "iconUri": "",
    "packageTelemetryId": "FLOW_GUID",
    "creator": "Vereniging Eigen Huis",
    "sourceEnvironment": ""
  },
  "resources": {
    "FLOW_GUID": {
      "type": "Microsoft.Flow/flows",
      "displayName": "VEH Copilot Training \u2014 Aanmelding",
      "description": "Ontvangt aanmeldingen van de online leeromgeving en stuurt een e-mailnotificatie naar de trainer.",
      "id": "/Microsoft.Flow/flows/FLOW_GUID",
      "registrationId": "FLOW_GUID",
      "isManaged": false,
      "details": {
        "displayName": "VEH Copilot Training \u2014 Aanmelding"
      },
      "configurableBy": "User",
      "hierarchy": "Root",
      "dependsOn": ["CONN_REG_ID"]
    },
    "CONN_REG_ID": {
      "type": "Microsoft.PowerApps/apis/connections",
      "displayName": "Office 365 Outlook",
      "description": "Microsoft Office 365 Outlook-verbinding voor het versturen van e-mails.",
      "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1636/1.0.1636.3409/office365/icon.png",
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
      "id": "/providers/Microsoft.PowerApps/apis/connections/shared_office365",
      "registrationId": "CONN_REG_ID",
      "isManaged": false,
      "details": {
        "displayName": "Office 365 Outlook"
      },
      "configurableBy": "User",
      "hierarchy": "Child",
      "dependsOn": []
    }
  }
}
'@
$manifest = $manifest.Replace("FLOW_GUID", $flowGuid).Replace("CONN_REG_ID", $connRegId)

# ── definition.json (de volledige flow-definitie) ─────────────────────────────
# Let op: $schema, $connections en $authentication zijn JSON-velden, geen PS-variabelen.
# De single-quoted here-string garandeert dat PS ze NIET expandeert.
$definition = @'
{
  "name": "FLOW_GUID",
  "properties": {
    "apiId": "/providers/Microsoft.PowerApps/apis/shared_logicflows",
    "displayName": "VEH Copilot Training \u2014 Aanmelding",
    "state": "Started",
    "connectionReferences": {
      "shared_office365": {
        "connectionName": "",
        "source": "Invoker",
        "id": "/providers/Microsoft.PowerApps/apis/shared_office365",
        "tier": "Standard"
      }
    },
    "definition": {
      "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
      "contentVersion": "1.0.0.0",
      "parameters": {
        "$connections":    { "defaultValue": {}, "type": "Object" },
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
              "Body": "<table style=\"font-family:Calibri,Arial,sans-serif;font-size:14px;border-collapse:collapse;width:100%;max-width:600px\"><tr style=\"background:#3B2785\"><td colspan=\"2\" style=\"padding:16px 20px;color:white;font-size:18px;font-weight:bold\">&#127891; Nieuwe aanmelding &#8212; Copilot Training VEH</td></tr><tr><td style=\"padding:10px 20px;width:140px;font-weight:bold;border-bottom:1px solid #eee\">Naam</td><td style=\"padding:10px 20px;border-bottom:1px solid #eee\">@{triggerBody()?['naam']}</td></tr><tr style=\"background:#f9f9f9\"><td style=\"padding:10px 20px;font-weight:bold;border-bottom:1px solid #eee\">Functie</td><td style=\"padding:10px 20px;border-bottom:1px solid #eee\">@{triggerBody()?['functie']}</td></tr><tr><td style=\"padding:10px 20px;font-weight:bold;border-bottom:1px solid #eee\">E-mailadres</td><td style=\"padding:10px 20px;border-bottom:1px solid #eee\">@{triggerBody()?['email']}</td></tr><tr style=\"background:#f9f9f9\"><td style=\"padding:10px 20px;font-weight:bold;border-bottom:1px solid #eee\">Leerwens</td><td style=\"padding:10px 20px;border-bottom:1px solid #eee\">@{triggerBody()?['leerwens']}</td></tr><tr><td style=\"padding:10px 20px;font-weight:bold\">Datum &amp; tijd</td><td style=\"padding:10px 20px\">@{triggerBody()?['datum']} om @{triggerBody()?['tijdstip']}</td></tr><tr style=\"background:#FFF3E0\"><td colspan=\"2\" style=\"padding:10px 20px;font-size:12px;color:#888\">Verstuurd vanuit de Online Leeromgeving Copilot in Dynamics 365 CE &bull; Vereniging Eigen Huis</td></tr></table>",
              "Importance": "Normal",
              "IsHtml": true
            }
          }
        },
        "Geef_HTTP_respons": {
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
$definition = $definition.Replace("FLOW_GUID", $flowGuid).Replace("TARGET_EMAIL", $Email)

# ── Schrijf bestanden ─────────────────────────────────────────────────────────
[System.IO.File]::WriteAllText("$tempDir\manifest.json", $manifest, $enc)
[System.IO.File]::WriteAllText("$flowDir\definition.json", $definition, $enc)

# ── Maak zip ──────────────────────────────────────────────────────────────────
if (Test-Path $outZip) { Remove-Item $outZip -Force }
Compress-Archive -Path "$tempDir\*" -DestinationPath $outZip -CompressionLevel Optimal
Remove-Item $tempDir -Recurse -Force

Write-Host ""
Write-Host "Klaar!" -ForegroundColor Green
Write-Host "  Bestand: $outZip" -ForegroundColor Cyan
Write-Host ""
Write-Host "Importeren in Power Automate:" -ForegroundColor Yellow
Write-Host "  1. Ga naar https://make.powerautomate.com" -ForegroundColor White
Write-Host "  2. Mijn flows > Importeren > Pakket (.zip) uploaden" -ForegroundColor White
Write-Host "  3. Koppel je Office 365 Outlook-verbinding" -ForegroundColor White
Write-Host "  4. Klik op 'Importeren'" -ForegroundColor White
Write-Host "  5. Open de flow en kopieer de HTTP POST-URL" -ForegroundColor White
Write-Host "  6. Plak de URL in index.html bij: const POWER_AUTOMATE_URL = ..." -ForegroundColor White
if ($Email -like "*VERVANG*") {
    Write-Host ""
    Write-Host "TIP: Geef je eigen e-mailadres mee:" -ForegroundColor Magenta
    Write-Host '  .\create_flow_package.ps1 -Email "jouw.naam@eigenhuis.nl"' -ForegroundColor Magenta
}
