function getGoogleLatLong([String]$address, [String]$city, [String]$country)
{
    $url = "https://maps.googleapis.com/maps/api/geocode/json?address=" + $address + "+" + $city + "+" + $country
    $result = Invoke-WebRequest -Uri $url
    $json = Invoke-WebRequest $url | ConvertFrom-JSON
    
    If(($json.status) -eq 'OVER_QUERY_LIMIT')
    {
        Write-Host $json.status
        Log-Write -LineValue "Processing has ended because we've exceeded Google's API limit. Try again in 24hrs."
        Log-Finish
        Exit
    }
    elseif(($json.results.geometry.location.lat) -eq '0')
    {
        Log-Write -LineValue "Received a 0 Lat. Aborted."
    }
    else {Return $json.results.geometry.location.lat, $json.results.geometry.location.lng}
}

function custAddressList
    {
        $sqlConnection = new-object System.Data.SqlClient.SqlConnection "Server=PFNZ-SRV-034;Database=CRM;Connection Timeout=600;Integrated Security=sspi"
        $sqlConnection.Open()
        $sqlCommand = $sqlConnection.CreateCommand()
        $query = "Select * from CRM.dbo.Customer_Address_Info ai`
                  where Ltrim(Rtrim(ai.[Physical Street #])) <> ''
                  and Ltrim(Rtrim(ai.[Physical City / Town])) <> ''
                  and Ltrim(Rtrim(ai.[Physical Street #])) <> ''
                  and ai.Branch = 'Maber Motors'"

        $sqlCommand.CommandText = $query
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlcommand
        $dataset = New-Object System.Data.DataSet
        $adapter.Fill($dataSet) | out-null
        $sqlConnection.Close()
        return $dataset
    }

function insertGeoMapping([string]$branch, [string]$CustId, [Float]$Lat, [Float]$Long, [String]$AddressStreet, [String]$AddressCity, [String]$AddressCountry)
{
    $insert = "INSERT INTO [dbo].[Customers_GeoMapping]
                ([Branch]
                ,[Customer ID]
                ,[Latitude]
                ,[Longitude]
                ,[DateLoaded]
                ,[AddressStreet]
                ,[AddressCity]
                ,[AddressCountry])
            VALUES
                ('" + $branch + "',
                '" + $CustId + "',
                " + $Lat + ",
                " + $Long + ",
                '" + $(Get-Date) + "',
                '" + $AddressStreet + "',
                '" + $AddressCity + "',
                '" + $AddressCountry + "'
                )"

    $sqlConnection = new-object System.Data.SqlClient.SqlConnection "Server=PFNZ-SRV-034;Database=CRM;Connection Timeout=600;Integrated Security=sspi"
    $sqlConnection.Open()
    $cmd = $sqlConnection.CreateCommand()
    $cmd.CommandText = $insert
    $rslt = $cmd.ExecuteNonQuery()
    $sqlconnection.Close()    
    return $rslt
}

function main
{
    $dsCA = New-Object System.Data.Dataset
    $dsCA = custAddressList    
    foreach($_ in $dsCA.tables[0].Rows)
    {
        $latlong = getGoogleLatLong($($_.'Physical Street #') + ' ' + $_.'Physical Street Name')($_.'Physical City / Town')('New Zealand')
        if(($latlong[0]) -ne '0')
        {
            $rslt = insertGeoMapping($($_.Branch))($($_.'ContactID'))($latlong[0])($latlong[1])($($_.'Physical Street #') + ' ' + `
                                    $_.'Physical Street Name')($_.'Physical City / Town')('New Zealand')
        }
        Start-Sleep -Seconds 1
    }        
}

Log-Start
."\\powerfarming.co.nz\netlogon\svn-netlogon\Logging_Functions.ps1"
    Main
Log-Finish

#getGoogleLatLong('19a Alpers Ridge','Cambridge', 'New Zealand')


