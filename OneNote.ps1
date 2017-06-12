
Function Get-OneNoteHeaders {

    [CmdletBinding()]

    Param
    ()

    Begin {
        $onenote = New-Object -ComObject OneNote.Application
        $scope = [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages
        [ref]$xml = $null
        $csvOutput = "onenote-headers.csv"
    }
    Process {     
        $onenote.GetHierarchy($null, $scope, $xml)
        [xml]$result = ($xml.Value)

        Foreach ($Notebook in $($result.DocumentElement.Notebook)) { 


            $sections = $Notebook | Select-Object -ExpandProperty section
            Foreach ($section in $sections) {
                
                $pages = $sections | Select-Object -ExpandProperty Page
                Foreach ($page in $pages) {
					$page | Format-List 
                 
					Write-Host -BackgroundColor White -ForegroundColor Black $( $page.ID + $page.name )
				    $onenote.GetPageContent($page.ID, $xml, [Microsoft.Office.Interop.OneNote.PageInfo]::piAll );
					
                    ########## At this point $xml.Value is the RAW TEXT of the page ##########
					# Write-Host -BackgroundColor Red -ForegroundColor White $xml.Value



					
					[xml]$result = ( Select-Xml -content $xml.Value -Xpath / )

                    [Reflection.Assembly]::LoadWithPartialName("System.Xml.Linq") | Out-Null
                    $XDocument = [System.Xml.Linq.XDocument]::Parse($xml.Value)

					format-list -InputObject $XDocument.Properties

                    foreach ($node in $nodes) {

                        Write-Host -BackgroundColor Magenta -ForegroundColor White $node
                        # if ($node.objectID != null)
                        # {
                        #     $node.objectID.Remove();
                        # } else if (node.Attribute("ID") != null)
                        # {
                        #     $node.id = "";
                        # }
                    }


                    # ConvertTo-Json -InputObject $XDocument. -Depth 12 | Set-Content -Path ( $( $page.name -replace '[^\w]+', '_' ) + ".json")

                    Set-Content -Value $XDocument -Path ( $( $page.name -replace '[^\w]+', '_' ) + ".xml")

					
	$node = $XDocument.Root.XPathSelectElement("one:OEChildren");
	Write-Host -BackgroundColor Cyan -ForegroundColor White $node

                    # Write-Host -BackgroundColor DarkBlue -ForegroundColor White $XDocument
					
					
					
					exit

				}
				
                Add-Content  $csvOutput "$($Notebook.name),$($section.name)"
            }
        }
    }
    End {
    }

}


<#
.DESCRIPTION


.INFO

.LINK XDOCUMENT api
https://msdn.microsoft.com/en-us/library/system.xml.linq.xdocument%28v=vs.110%29.aspx?f=255&MSPPError=-2147217396
#>
function Get-OneDoc {
	param(
		[string] $PageId
	)
	$node = [System.Xml.Linq.XDocument]::XPathSelectElement($XDocument,"one:OEChildren");
	Write-Host -BackgroundColor Cyan -ForegroundColor White $node
	

}



Get-OneNoteHeaders



