        



#region F2 Image Extraction

Function FtMain
{

    $wc = New-Object System.Net.WebClient
    Log-Write -LineValue 'F2MAIN' 

    $branchList = @{'Waikato' = 'Maber Motors';`
                    'ashburton' = 'Ashburton';`
                    'canterbury' = 'Canterbury';`
                    'gisborne' = 'Gisborne';`
                    'gore' = 'Gore';`
                    'hawkesbay' = 'Hawkes Bay';`
                    'invercargill' = 'Invercargill';`
                    'manawatu' = 'Manawatu';`
                    'northland' = 'Northland';`
                    'otago' = 'Otago';`
                    'taranaki' = 'Taranaki';`
                    'teawamutu' = 'Te Awamutu';
                    'timaru' = 'Timaru';
                    'westcoast' = 'WestCoast'}
                                        
    foreach($a In $branchList.Keys)
    {
        InitFtBranchFolders($($branchlist.Item($a)))
        If(Test-Path -Path "$ftStagingFolder\$($branchlist.Item($a))\${a}.xml")
            {Remove-Item "$ftStagingFolder\$($branchlist.Item($a))\${a}.xml"}        
        Try
        {
            $wc.DownloadFile("$ftUrl$a.xml", "$ftStagingFolder\$($branchlist.Item($a))\${a}.xml")
            downloadFtListings("$ftStagingFolder\$($branchlist.Item($a))\${a}.xml")($a)($($branchlist.Item($a)))
        }
        Catch{Log-Write -LineValue "Failed to download $ftUrl$a.xml"}
    }
}    

Function compareAxToUpdated
{
    Log-Write -LineValue "compareAxToUpdated:"

    # Loop through AX folders
    foreach ($folder in (Get-ChildItem -Path "$axStagingFolder" -Directory))
    {
        # Loop through files in TEMP
        foreach ($file in (Get-ChildItem -Path "$($folder.FullName)\Temp"))
        {
            # Is there a Base, or is this image new.
            If(Test-Path -Path "$($folder.FullName)\$($file.Name)")
            {
                # Compare file size of download with Base.
                If($file.Length -ne (Get-Item "$($folder.FullName)\$($file.Name)").Length)
                {
                    # Updated file so copy to Base.
                    Copy-Item $file.FullName $folder.FullName -Force

                    # If the image exists in the FastTrack base folder, abandon update.
                    If(!(Test-Path "$($ftStagingFolder)\$($folder.BaseName)\$($file.Name)"))
                        {                            
                            # Use AX image, but compare file sizes for update first.
                            If($file.Length -ne (Get-Item "$($updatedFinalFolder)\$($file.Name)").Length)
                            {
                                Copy-Item $file.FullName $updatedFinalFolder -Force
                                Log-Write -LineValue "  New Ax File $($folder.FullName)\Temp\$($file.Name) found and copied to Main Update folder."
                            }
                            else
                            {
                                Log-Write -LineValue "  Ax file $($folder.FullName)\Temp\$($file.Name) is indentical."
                            }
                        }
                        else
                        {
                            Log-Write -LineValue "  File $($folder.FullName)\Temp\$($file.Name) skipped compare because FastTrack image exists." 
                        }
                    Remove-Item -Path $file.FullName -Force                    
                }
                else
                {
                    # No Update Needed
                    Remove-Item -Path $file.FullName -Force
                    Log-Write -LineValue "  No Update Needed. File size for $($file.FullName) is identical."
                }                
            }
            else
            {
                # New File
                Copy-Item $file.FullName $folder.FullName
                If(!(Test-Path "$($updatedFinalFolder)\$($file.Name)"))
                    {
                        Copy-Item $file.FullName $updatedFinalFolder -Force
                        Log-Write -LineValue "  Newer File $($folder.FullName)\Temp\$($file.Name) found and copied to Main Update folder."                         
                    }
                    else
                    {
                        Log-Write -LineValue "  Newer File $($folder.FullName)\Temp\$($file.Name) found but skipped because FastTrack image exists." 
                    }
                Remove-Item -Path $file.FullName
            }
        }
    }    
}

Function compareFtToUpdated
{
    Log-Write -LineValue "compareFtToUpdated:"
    foreach ($folder in (Get-ChildItem -Path "$ftStagingFolder" -Directory))
    {
        foreach ($file in (Get-ChildItem -Path "$($folder.FullName)\Temp"))
        {
            If(Test-Path -Path "$($folder.FullName)\$($file.Name)")
            {
                If($file.Length -ne (Get-Item "$($folder.FullName)\$($file.Name)").Length)
                {
                    # Updated File
                    Copy-Item $file.FullName $folder.FullName -Force
                    #Copy-Item $file.FullName "$($folder.FullName)\Updated" -Force
                    Copy-Item $file.FullName $updatedFinalFolder -Force                                                
                    Remove-Item -Path $file.FullName -Force                    
                }
                else
                {
                    # No Update Needed
                    Remove-Item -Path $file.FullName -Force
                    Log-Write -LineValue "  No Update Needed. File size for $($file.FullName) is identical."
                }                
            }
            else
            {
                # New File
                Copy-Item $file.FullName $folder.FullName
                #Copy-Item $file.FullName "$($folder.FullName)\Updated"
                Copy-Item $file.FullName $updatedFinalFolder -Force
                Remove-Item -Path $file.FullName
                Log-Write -LineValue "  New File $($folder.FullName)\Temp\$($file.Name). Added to Base and Updated Folders."
            }
        }
    }    
}

Function downloadFtListings([String]$xmlFt, [String]$ftBranch, [String]$stgBranch)
{

    # Load Up DS with Veh Listing for Validation
    $dsF2v = New-Object System.Data.Dataset
    $dsF2v = F2Vehicles

    # Iterate Listings in XML and Download to Temp folders
    $wc2 = New-Object System.Net.WebClient
    [xml]$XmlDocument = Get-Content -Path $xmlFt
    foreach ($listing in $XmlDocument.dealer.listing)
    {
        # F2 StockNumbers are Int datatypes. Skip else.
        If([Microsoft.VisualBasic.Information]::IsNumeric($($listing.stock_number)))
        {
            # Qeury F2 for Exact Stock# Match in FastTrack
            $veh = $dsF2v.tables[0].select("[Stock #] = '$($listing.stock_number)' and [Branch] = '$($stgBranch)'")
            If($veh.Length -ne 0)
            {
                Try
                {
                $wc2.DownloadFile("$($ftUrl)$($ftBranch)_$($listing.id)_1.jpg", `
                    "$ftStagingFolder\$($stgBranch)\Temp\$($stgBranch)_$($listing.stock_number).jpg")
                Log-Write -LineValue "  Downloaded: $($ftUrl)$($ftBranch)_$($listing.id)_1.jpg to `
                    $ftStagingFolder\$($stgBranch)\Temp\$($stgBranch)_$($listing.stock_number).jpg"
                }
                Catch{Log-Write -LineValue "Failed to download $($ftUrl)$($ftBranch)_$($listing.id)_1.jpg"}
            }
            else 
            { 
                Log-Write -LineValue "  Skipped: $($ftUrl)$($ftBranch)_$($listing.id)_1.jpg"
            }
        }
    }
}

Function F2Vehicles
    {
	    $query = ""
        $query = "SELECT [Branch],[Stock #] FROM [CRM].[dbo].[Vehicle_XML]"

        $sqlConnection = new-object System.Data.SqlClient.SqlConnection "Server=PFNZ-SRV-034;Database=CRM;Connection Timeout=600;Integrated Security=sspi"
        $sqlConnection.Open()

        #Create a command object
        $sqlCommand = $sqlConnection.CreateCommand()
        $sqlCommand.CommandText = $query

        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlcommand
        $dataset = New-Object System.Data.DataSet

        $adapter.Fill($dataSet) | out-null
        $sqlConnection.Close()
        return $dataset
    }


Function InitFtBranchFolders([String]$Branch)
{    
    If (!(Test-Path -Path "$ftStagingFolder\$Branch"))
        {New-Item -Path "$ftStagingFolder\$Branch" -ItemType Directory | Out-Null
            Log-Write -LineValue "Created folder $ftStagingFolder\$Branch"}

    If (!(Test-Path "$ftStagingFolder\$Branch\Temp"))
        {New-Item -Path "$ftStagingFolder\$Branch\Temp" -ItemType Directory | Out-Null
            Log-Write -LineValue "Created folder $ftStagingFolder\$Branch\Temp"}

    If (!(Test-Path "$ftStagingFolder\$Branch\Updated"))
        {New-Item -Path "$ftStagingFolder\$Branch\Updated" -ItemType Directory | Out-Null
            Log-Write -LineValue "Created folder $ftStagingFolder\$Branch\Updated"}
}


#endregion

#region AX Image Extraction

Function AXMain
{
    Log-Write -LineValue 'AX Image Processing Begins'

    #Loop Branches / Create Folder Structure            
    $dsBL = New-Object System.Data.Dataset
    $dsBL = VehicleBranchList
    foreach($a in $dsBL.tables[0].Rows)
    {                
        #ExtractAxItemImageToFile("c:\support\test\asdf.png")($imgdata)
        InitAxBranchFolders($a.Branch)
    }    

    #Loop Vehicles with AX Images / Extract to Temp               
    $dsAV = New-Object System.Data.Dataset
    $dsAV = AXVehicles    
    foreach($a in $dsAV.tables[0].Rows)
    {        
        [Byte[]]$imgdata = $($a.Koo_Image)
        ExtractAxItemImageToFile("$axStagingFolder\$($a.Branch)\Temp\$($a.Branch)_$($a.'Stock #').$($a.'KOO_IMAGEFORMAT')")($imgdata) | Out-Null                        
        If (Test-Path "$axStagingFolder\$($a.Branch)\Temp\$($a.Branch)_$($a.'Stock #').$($a.'KOO_IMAGEFORMAT')")
            {Log-Write -LineValue "  Successfully extracted $axStagingFolder\$($a.Branch)\Temp\$($a.Branch)_$($a.'Stock #').$($a.'KOO_IMAGEFORMAT')"}
            Else
            {Log-Write -LineValue "  Failed to extract $axStagingFolder\$($a.Branch)\Temp\$($a.Branch)_$($a.'Stock #').$($a.'KOO_IMAGEFORMAT')"}
    }                                
}

Function InitAxBranchFolders([String]$Branch)
{    
    If (!(Test-Path -Path "$axStagingFolder\$Branch"))
        {New-Item -Path "$axStagingFolder\$Branch" -ItemType Directory | Out-Null
            Log-Write -LineValue "Created folder $axStagingFolder\$Branch"}

    If (!(Test-Path "$axStagingFolder\$Branch\Temp"))
        {New-Item -Path "$axStagingFolder\$Branch\Temp" -ItemType Directory | Out-Null
            Log-Write -LineValue "Created folder $axStagingFolder\$Branch\Temp"}

    If (!(Test-Path "$axStagingFolder\$Branch\Updated"))
        {New-Item -Path "$axStagingFolder\$Branch\Updated" -ItemType Directory | Out-Null
            Log-Write -LineValue "Created folder $axStagingFolder\$Branch\Updated"}
}

Function AXVehicles
    {
	    $query = ""
        $query = "select vi.*, it.KOO_IMAGE, it.KOO_IMAGEFORMAT "
	    $query = $query + "from [CRM].[dbo].[Vehicle_Info] vi inner join"
        $query = $query + "[PFWAX].[PFW_AX2009_Live].[dbo].[INVENTTABLE] it "
        $query = $query + "on it.DATAAREAID = 'pfw' "
        $query = $query + "and it.ITEMID COLLATE DATABASE_DEFAULT = vi.[Item ID] COLLATE DATABASE_DEFAULT "
        $query = $query + "where it.KOO_IMAGE is not null "
        $query = $query + "and it.KOO_IMAGEFORMAT <> ''"

        $sqlConnection = new-object System.Data.SqlClient.SqlConnection "Server=PFNZ-SRV-034;Database=CRM;Connection Timeout=600;Integrated Security=sspi"
        $sqlConnection.Open()

        #Create a command object
        $sqlCommand = $sqlConnection.CreateCommand()
        $sqlCommand.CommandText = $query

        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlcommand
        $dataset = New-Object System.Data.DataSet

        $adapter.Fill($dataSet) | out-null
        $sqlConnection.Close()
        return $dataset
    }

Function ExtractAxItemImageToFile([String]$pathToFile, [byte[]]$flBinary)
    {        
        #$conn = New-Object -comobject ADODB.Connection
        $stream2 = new-object -comobject ADODB.Stream
        $stream2.Type = 1
        $stream2.Open()

		$stream = new-object -comobject ADODB.Stream
		$stream.Type = 1
		$stream.Open()	
		$stream.Write($flBinary)
		$stream.Position = 7
		$stream.CopyTo($stream2, ($stream.size - 7))
		$stream2.SaveToFile($pathToFile, 2)
        $stream.close
        $stream2.close
    }

Function VehicleBranchList
    {
        $sqlConnection = new-object System.Data.SqlClient.SqlConnection "Server=PFNZ-SRV-034;Database=CRM;Connection Timeout=600;Integrated Security=sspi"
        $sqlConnection.Open()

        #Create a command object
        $sqlCommand = $sqlConnection.CreateCommand()
        $query = ""
        $query = $query + "SELECT Distinct [Branch] "
        $query = $query + "FROM [CRM].[dbo].[Vehicle_XML];"
        $sqlCommand.CommandText = $query

        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlcommand
        $dataset = New-Object System.Data.DataSet

        $adapter.Fill($dataSet) | out-null

        #write-output $dataset.Tables[0].Rows.Count

        # Close the database connection
        $sqlConnection.Close()
        return $dataset
    }

#endregion

Add-Type -Assembly Microsoft.VisualBasic

."\\powerfarming.co.nz\netlogon\svn-netlogon\Logging_Functions.ps1"
Log-Start

[String]$axStagingFolder = "C:\ExtractUploadVehicleImagesForCRM\AX"    
[String]$ftStagingFolder = "C:\ExtractUploadVehicleImagesForCRM\FT"  
[String]$updatedFinalFolder = "C:\ExtractUploadVehicleImagesForCRM\Updates"
[String]$ftUrl = "http://usedproductws-prod.powerfarming.co.nz/data/"

    # Create Folders
    If(!(Test-Path -Path $axStagingFolder)){New-Item -Path $axStagingFolder -ItemType Directory }
    If(!(Test-Path -Path $ftStagingFolder)){New-Item -Path $ftStagingFolder -ItemType Directory}
    If(!(Test-Path -Path $updatedFinalFolder)){New-Item -Path $updatedFinalFolder -ItemType Directory}

FtMain
compareFtToUpdated
AXMain
compareAxToUpdated

Log-Finish



