#
# PowerFarming.CRM
#
function Get-CrmGeoMapBackLogData([string]$mssqldbserver, [string]$dbname, [string]$username, [string]$password)
    {
        $sqlConnection = new-object System.Data.SqlClient.SqlConnection "Server=$mssqldbserver;Database=$dbname;Connection Timeout=600;User Id=$username; Password=$password"        
        $sqlConnection.Open()
        $sqlCommand = $sqlConnection.CreateCommand()
        
        $query = "SELECT Count(*) as ToGeoMap, Max(datediff(dd,getdate(),DateLoaded)*-1) as AgeDays
                    FROM [CRM].[dbo].[Customers_GeoMapping]
                    where Latitude is null
                    and Error <> 1
                    group by [Loaded to CRM]"

        $sqlCommand.CommandText = $query
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlcommand
        $dataset = New-Object System.Data.DataSet
        $adapter.Fill($dataSet) | out-null
        $sqlConnection.Close()
        
        foreach ($_ in $dataset.Tables[0].Rows)
        {
            $ar = @{"ToGeoMap" = $_.ToGeoMap; "AgeDays" = $_.AgeDays}
            return $ar
        }
                
    }

<#
.SYNOPSIS
 To return the progress ( or lack thereof ) of queued un-calculated LatLong records
 used by CRM from F2. 

.DESCRIPTION
 We queue up F2 customer records and their addresses and update LatLongs to be later 
 loaded or upserted into CRM using SSIS and KingswaySoft CRM integration.
 This function returns XML formatted as a PRTG advanced customer sensor to track the 
 backlog of outstanding queue records.

.PARAMETER mssqldbserver
 The MS SQL server and instance ( optional ) where the database and table resides.
 Assumes SSPI authentication.

.PARAMETER dbname
 The name of the database hosting the table to update.
 The table name is hardcoded as dbo.Customers_GeoMapping

.PARAMETER Force
 If an existing object is not found, instead of writing an error, a new
 instance of the object will be created and returned.

.EXAMPLE
 Get-PrtgXmlCrmGeoMapBackLog("pfnz-srv-034.powerfarming.co.nz", "crm")

#>

function Get-PrtgXmlCrmGeoMapBackLog
    {

		Param(
			[Parameter(Mandatory=$true)][string] $mssqldbserver, 
			[Parameter(Mandatory=$true)][string] $dbname,
            [Parameter(Mandatory=$true)][string] $username,
            [Parameter(Mandatory=$true)][string] $password)

        $ar = Get-CrmGeoMapBackLogData($mssqldbserver)($dbname)($username)($password)
        $tmpfile = "$env:temp\PrtgXmlCrmGeoMapBackLog-$(Get-Date -Format 'yyyy-MM-dd-hhmmss').xml"
		$xmlWr = New-Object System.XMl.XmlTextWriter($tmpfile,$null)	
		$xmlWr.Formatting = 'Indented'
		$xmlWr.Indentation = 1
		$xmlWr.IndentChar = "`t"
		$xmlWr.WriteStartDocument()
		$xmlWr.WriteStartElement('prtg')

		$xmlWr.WriteStartElement('result')    
        $xmlWr.WriteElementString('channel','ToGeoMap')
        $xmlWr.WriteElementString('value',$($AR.Get_Item("ToGeoMap")))
        $xmlWr.WriteEndElement()
        
        $xmlWr.WriteStartElement('result')    
        $xmlWr.WriteElementString('channel','AgeDays')
        $xmlWr.WriteElementString('value',$($AR.Get_Item("AgeDays")))
        $xmlWr.WriteEndElement()
        
        $xmlWr.WriteEndElement()
		$xmlWr.WriteEndDocument()
		$xmlWr.Flush()
		$xmlWr.Close()

		$prtgxml = Get-Content $tmpfile
		Remove-Item -Path $tmpfile
		return $prtgxml

    }
export-modulemember -Function Get-PrtgXmlCrmGeoMapBackLog









