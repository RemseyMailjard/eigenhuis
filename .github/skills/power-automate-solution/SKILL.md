---
name: power-automate-solution
description: >
  Use this skill whenever you need to generate an importable Power Automate
  solution package (.zip) via PowerShell. Covers creating the correct Dataverse
  solution structure (solution.xml, customizations.xml, [Content_Types].xml,
  Workflows/*.json) so the zip imports without errors into
  make.powerautomate.com > Solutions > Import Solution.
  Trigger on keywords: Power Automate, flow, solution zip, import flow,
  ModernFlow, HTTP trigger, aanmelding, workflow package.
---

# Power Automate Solution Skill

## Doel

Genereer een importeerbaar Power Automate-oplossingspakket (.zip) volledig via
PowerShell — geen handmatige stappen in de browser vereist.

---

## Vereiste bestandsstructuur in de zip

```
[Content_Types].xml      ← verplicht, exact deze naam
solution.xml             ← verplicht
customizations.xml       ← verplicht
Workflows/
  {FlowName}-{GUID-UPPERCASE}.json
```

> ⚠️ De zip mag ALLEEN bestanden bevatten. Geen losse mappen als root-entries.

---

## Kritieke regels (geleerd uit mislukte imports)

### 1. GUID-schrijfwijze

| Locatie                             | Schrijfwijze                                      |
| ----------------------------------- | ------------------------------------------------- |
| `solution.xml` `RootComponent id=`  | **lowercase** `{a1b2c3d4-...}`                    |
| `customizations.xml` `WorkflowId=`  | **lowercase** `{a1b2c3d4-...}`                    |
| `customizations.xml` `JsonFileName` | **UPPERCASE** `/Workflows/Name-A1B2C3D4-....json` |
| Bestandsnaam in Workflows/          | **UPPERCASE** `Name-A1B2C3D4-....json`            |

Dataverse vergelijkt het RootComponent-id **case-sensitive** met de WorkflowId
in de flow runtime. Mismatch → fout `0x8004803A`.

### 2. solution.xml — RootComponents

**Alleen** de workflow (type 29) declareren via `id=`. De connection reference
(type 10466) hoort NIET in RootComponents:

```xml
<RootComponents>
  <RootComponent type="29" id="{a1b2c3d4-e5f6-7890-abcd-ef1234567890}" behavior="0" />
</RootComponents>
```

### 3. customizations.xml — Workflow-element

Gebruik **child-elementen** (niet attributen) voor StateCode/StatusCode.
Verplichte velden die bij een echte export aanwezig zijn:

```xml
<Workflow WorkflowId="{lowercase-guid}" Name="Naam van de flow">
  <JsonFileName>/Workflows/Naam-UPPERCASE-GUID.json</JsonFileName>
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
    <LocalizedName languagecode="1033" description="Naam van de flow" />
  </LocalizedNames>
</Workflow>
```

Verplichte lege secties in customizations.xml:

```xml
<Entities></Entities>
<Roles></Roles>
<FieldSecurityProfiles></FieldSecurityProfiles>
<Templates />
<EntityMaps />
<EntityRelationships />
<OrganizationSettings />
<optionsets />
<CustomControls />
<EntityDataProviders />
<Languages><Language>1033</Language></Languages>
```

### 4. Workflow JSON — OpenApiConnection (niet ApiConnection)

Gebruik het formaat dat een echte export produceert:

```json
{
  "properties": {
    "connectionReferences": {
      "shared_office365": {
        "api": { "name": "shared_office365" },
        "connection": { "connectionReferenceLogicalName": "new_logical_name" },
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
              "emailMessage/To": "ontvanger@domein.nl",
              "emailMessage/Subject": "Onderwerp",
              "emailMessage/Body": "<html>...</html>",
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
            "schema": { "type": "object", "properties": { ... } }
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
```

> ⚠️ Gebruik NOOIT `"type": "ApiConnection"` met een `"body"`-structuur —
> dat is het oude formaat en wordt afgewezen.

---

## PowerShell-patroon voor zip aanmaken

```powershell
$enc    = [System.Text.UTF8Encoding]::new($false)   # UTF-8 zonder BOM
$tempDir = Join-Path $env:TEMP "FlowPkg_$(Get-Random)"
New-Item -ItemType Directory -Force -Path "$tempDir\Workflows" | Out-Null

# schrijf bestanden
[System.IO.File]::WriteAllText("$tempDir\[Content_Types].xml",  $contentTypes,  $enc)
[System.IO.File]::WriteAllText("$tempDir\solution.xml",         $solutionXml,   $enc)
[System.IO.File]::WriteAllText("$tempDir\customizations.xml",   $custXml,       $enc)
[System.IO.File]::WriteAllText("$tempDir\Workflows\$filename",  $flowJson,      $enc)

# zip
Compress-Archive -Path "$tempDir\*" -DestinationPath $outZip -CompressionLevel Optimal
Remove-Item $tempDir -Recurse -Force
```

> Gebruik altijd `[System.IO.File]::WriteAllText` met expliciete UTF-8 encoder,
> nooit `Set-Content` (kan BOM of encoding-issues veroorzaken).

---

## Verificatie na aanmaken

```powershell
# Controleer inhoud van de zip
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($outZip)
$zip.Entries | Select-Object FullName
$zip.Dispose()

# Controleer RootComponent id in solution.xml
$sol = ... # lees solution.xml uit zip
$xml.ImportExportXml.SolutionManifest.RootComponents.RootComponent | Format-List
```

Verwachte output:

```
type     : 29
id       : {a1b2c3d4-e5f6-7890-abcd-ef1234567890}   ← lowercase
behavior : 0
```

---

## Importeren in Power Automate

1. Ga naar **make.powerautomate.com**
2. Klik links op **Solutions** (Oplossingen)
3. Klik op **Import Solution → Browse** → kies de `.zip`
4. Klik op **Next** → koppel de **Office 365 Outlook**-verbinding → **Import**
5. Open de geïmporteerde flow → kopieer de **HTTP POST-URL** uit de trigger
6. Gebruik de URL als endpoint in de aanroepende applicatie

> ⚠️ Gebruik **Solutions → Import Solution**, niet "Mijn flows → Importeren →
> Pakket (.zip)" — dat is het oude (niet meer ondersteunde) formaat.

---

## Referentie: demo-export vergelijken

Als een import blijft mislukken:

1. Maak een simpele testflow handmatig in Power Automate
2. Exporteer via **Solutions → Export** als unmanaged
3. Pak de zip uit en vergelijk `solution.xml`, `customizations.xml` en het
   workflow-JSON met jouw gegenereerde versie
4. Let specifiek op: GUID-casing, aanwezige XML-elementen, JSON `type`-veld

Dit "reverse engineering"-patroon is de meest betrouwbare manier om
structuurfouten te vinden.
