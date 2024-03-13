# Obtain Postman API Key, if cannot be found request the user provides one
if (-not [string]::IsNullOrEmpty($env:POSTMAN_API_KEY)) {
    $postmanApiKey = $env:POSTMAN_API_KEY;
} 
else {
    $secureapiKey = Read-Host "Please enter your Postman API Key (the one from your Postman account settings, e.g. PMAK-xxxxxxxxxxxxxxxxxxxxxxxx-XXXX)" -AsSecureString;
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureapiKey);
    $postmanApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR);
}

# Obtain Payroc API Key, if cannot be found request the user provides one
if (-not [string]::IsNullOrEmpty($env:PAYROC_API_KEY)) {
    $payrocApiKey = $env:PAYROC_API_KEY;
} 
else {
    $secureapiKey = Read-Host "Please enter your Payroc API Key (the one provided to you by Payroc)" -AsSecureString;
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureapiKey);
    $payrocApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR);
}

# Write Payroc Key into Environment file
$envPath = ".\environments\payroc-uat.postman_environment.json";
(Get-Content -Path $envPath -Raw) -replace '{{PayrocApiKey}}', $payrocApiKey | Out-File -FilePath $envPath

# Set the Postman API URL
$postmanApiUrl = "https://api.postman.com";

# Set the headers for authentication
$headers = @{
    "X-Api-Key" = $postmanApiKey
    "Content-Type" = "application/json"
};

$workspaceName = "Payroc API";

# Step 1: Create Workspace if it does not exist and capture workspace Id
$existingWorkSpaces = Invoke-RestMethod -Method Get -Uri "$postmanApiUrl/workspaces" -Headers $headers;

# Capture existing workspace if it exists
foreach ($existingWorkSpace in $existingWorkSpaces.workspaces) {
    if ($existingWorkSpace.name -eq $workspaceName) {
        $workSpaceId = $existingWorkSpace.id;

        Write-Host "Identified Workspace '$workspaceName' with Id $workSpaceId";
    }
}

# If workspace does not exist, create the workspace
if ([String]::IsNullOrEmpty($workSpaceId)) {
    $workSpaceData = @{
        name = $workspaceName;
        type = "personal";
        description = "Workspace for $workspaceName";
    };
    
    $workspace = @{
        workspace = $workspaceData;
    } | ConvertTo-Json;

    Write-Host "Creating Workspace '$workspaceName'...";

    $newWorkspace = Invoke-RestMethod -Uri "$postmanApiUrl/workspaces" -Headers $headers -Method Post -Body $workspace;

    $workSpaceId = $newWorkspace.workspace.id;
    
    Write-Host "New Workspace '$workspaceName' has Id $workSpaceId";
}

Write-Host "Updating Workspace $workspaceName";

$workSpaceQuery="?workspace=$workSpaceId";

$environmentFolderPath = "$PSScriptRoot\environments";

# Get a list of JSON files in the specified folder
$environmentFiles = Get-ChildItem -Path $environmentFolderPath -Filter "*.json";

$existingEnvironments = Invoke-RestMethod -Method Get -Uri "$postmanApiUrl/environments$workSpaceQuery" -Headers $headers;

# Step 2: Import Environments
foreach ($environmentFile in $environmentFiles) {
    $environment = Get-Content -Path $environmentFile.FullName;

    $environmentData = $environment | ConvertFrom-Json;
    
    # Step 3: Delete environment if we intend to import it
    foreach ($existingEnvironment in $existingEnvironments.environments) {
        if ($existingEnvironment.name -eq $environmentData.name) {
            $resourceId = $existingEnvironment.id;
        
            $environmentName = $existingEnvironment.name;
        
            Write-Host "Deleting environment '$environmentName'...";
            Invoke-RestMethod -Uri "$postmanApiUrl/environments/$resourceId" -Headers $headers -Method Delete | Out-Null;
        }
    }

    $requestBody = @{
        environment = $environmentData
    } | ConvertTo-Json -Depth 100;

    $environmentName = $environmentData.name;

    # Step 4: Create environment
    Write-Host "Creating environment '$environmentName'...";
    Invoke-RestMethod -Uri "$postmanApiUrl/environments$workSpaceQuery" -Headers $headers -Method Post -Body $requestBody | Out-Null;
}

Write-Host "Environments imported.";

# Specify the folder path containing the collection JSON files

$collectionFolderPath = "$PSScriptRoot\collections";

# Get a list of JSON files in the specified folder
$collectionFiles = Get-ChildItem -Path $collectionFolderPath -Filter "*.json";

$existingCollections = Invoke-RestMethod -Method Get -Uri "$postmanApiUrl/collections$workSpaceQuery" -Headers $headers;

# Step 5: Import Collections
foreach ($collectionFile in $collectionFiles) {
    $collection = Get-Content -Path $collectionFile.FullName;

    $collectionData = $collection | ConvertFrom-Json;

    # Step 6: Delete collection if we intend to import it
    foreach ($existingCollection in $existingCollections.collections) {
        if ($existingCollection.name -eq $collectionData.info.name) {
            $resourceId = $existingCollection.id;
        
            $collectionName = $existingCollection.name;
        
            Write-Host "Deleting collection '$collectionName'...";
            Invoke-RestMethod -Uri "$postmanApiUrl/collections/$resourceId" -Headers $headers -Method Delete | Out-Null;
        }
    }

    $requestBody = @{
        collection = $collectionData
    } | ConvertTo-Json -Depth 100;

    $collectionName = $collectionData.info.name;

    # Step 7: Create collection
    Write-Host "Creating collection '$collectionName'...";
    Invoke-RestMethod -Uri "$postmanApiUrl/collections$workSpaceQuery" -Headers $headers -Method Post -Body $requestBody | Out-Null;
}

Write-Host "Collections imported.";
Write-Host "Setup complete.";
Write-Host;
