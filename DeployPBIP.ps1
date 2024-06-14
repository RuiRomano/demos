# Description: This script is used to deploy a PBIP to a Fabric workspace, and manipulate the PBIP file before publishing it.
# Learn more: https://learn.microsoft.com/en-us/rest/api/fabric/articles/get-started/deploy-project

$ErrorActionPreference = "Stop"

# Parameters 

$workspaceName = "RR - PBIR - Demo 01"
$pbipSemanticModelPath = ".\PBIP\Sales.SemanticModel"
$pbipReportPath = ".\PBIP\Sales.Report"
$environment = "TST"

$currentPath = (Split-Path $MyInvocation.MyCommand.Definition -Parent)
Set-Location $currentPath

# Download modules and install

New-Item -ItemType Directory -Path ".\modules" -ErrorAction SilentlyContinue | Out-Null
@("https://raw.githubusercontent.com/microsoft/Analysis-Services/master/pbidevmode/fabricps-pbip/FabricPS-PBIP.psm1"
, "https://raw.githubusercontent.com/microsoft/Analysis-Services/master/pbidevmode/fabricps-pbip/FabricPS-PBIP.psd1") |% {
    Invoke-WebRequest -Uri $_ -OutFile ".\modules\$(Split-Path $_ -Leaf)"
}
if(-not (Get-Module Az.Accounts -ListAvailable)) { 
    Install-Module Az.Accounts -Scope CurrentUser -Force
}

Import-Module ".\modules\FabricPS-PBIP" -Force

# Authenticate

Set-FabricAuthToken

# Ensure workspace exists

$workspaceId = New-FabricWorkspace  -name $workspaceName -skipErrorIfExists

# Import the semantic model into Fabric workspace

$semanticModelImport = Import-FabricItem -workspaceId $workspaceId -path $pbipSemanticModelPath
#$semanticModelImport = @{ "Id" = "4db8f21a-ed48-4328-abe1-4c1eb43438c3" }

# Manipulate the PBIR before publishing

$definitionPath = "$pbipReportPath\definition"

## Set default page

$json = Get-Content "$definitionPath\pages\pages.json" | ConvertFrom-Json

$json.activePageName = "c2d9b4b1487b2eb30e98"

$json | ConvertTo-Json -Depth 100 | Set-Content "$definitionPath\pages\pages.json"

## Set default values on slicers

$slicerFiles = Get-ChildItem $definitionPath -Recurse -Filter "visual.json" | Where-Object { 
    $json = Get-Content $_.FullName | ConvertFrom-Json
    
    $json.name -in @("4e0e638edcb47268821e", "f42f1648ed721ff28383")
 }

foreach ($file in $slicerFiles) {

    $json = Get-Content $file.FullName | ConvertFrom-Json

    # Year slicer
    if ($json.name -eq "4e0e638edcb47268821e")
    {
        $yearFilterJson = "        
            {""filter"": {
                ""Version"": 2,
                ""From"": [
                  {
                    ""Name"": ""c"",
                    ""Entity"": ""Calendar"",
                    ""Type"": 0
                  }
                ],
                ""Where"": [
                  {
                    ""Condition"": {
                      ""In"": {
                        ""Expressions"": [
                          {
                            ""Column"": {
                              ""Expression"": {
                                ""SourceRef"": {
                                  ""Source"": ""c""
                                }
                              },
                              ""Property"": ""Year""
                            }
                          }
                        ],
                        ""Values"": [
                          [
                            {
                              ""Literal"": {
                                ""Value"": ""$([datetime]::Now.Year)L""
                              }
                            }
                          ]
                        ]
                      }
                    }
                  }
                ]
              }
            }"
                
        if ($json.visual.objects.general.properties.filter)
        {
            $json.visual.objects.general.properties.PSObject.Properties.Remove('filter')
        }

        $json.visual.objects.general.properties | Add-Member -MemberType NoteProperty -Name "filter" -Value ($yearFilterJson | ConvertFrom-Json) -Force

    }
    # Country slicer - remove any filter
    elseif ($json.name -eq "f42f1648ed721ff28383")
    {        
        if ($json.visual.objects.general.properties.filter)
        {
            $json.visual.objects.general.properties.PSObject.Properties.Remove('filter')
        }
    }

    $json | ConvertTo-Json -Depth 100 | Set-Content $file.FullName
}

## Add watermark visual into every page, copying the visual from a folder

$visualPath = ".\_teamVisuals\watermark"

$pageFolders = Get-ChildItem "$definitionPath\pages" -Directory

foreach ($pageFolder in $pageFolders) {

    Copy-Item $visualPath -Recurse -Destination "$($pageFolder.FullName)\visuals" -Force
    
    # Replace placeholder text of the watermark visual

    $newContent = Get-Content "$($pageFolder.FullName)\visuals\watermark\visual.json" |% { $_ -replace "\[placeholdertext\]","$environment - $([datetime]::Now.ToString("s"))"  }

    $newContent | Out-File "$($pageFolder.FullName)\visuals\watermark\visual.json" -Force
}

# Import the report into Fabric workspace

$reportImport = Import-FabricItem -workspaceId $workspaceId -path $pbipReportPath -itemProperties @{"semanticModelId" = $semanticModelImport.Id}

