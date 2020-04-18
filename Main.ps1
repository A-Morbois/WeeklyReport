param (
    [Switch] $DebugSoft
)


#region Tool
function Azure_CreateHeader
{

param(
    [String] $_appType
 )

    $Auth = '{0}:{1}' -f $Azure_Username, $Azure_Password;
    $Auth = [System.Text.Encoding]::UTF8.GetBytes($Auth);
    $Auth = [System.Convert]::ToBase64String($Auth);

    $HeaderforAzure = 
    @{
        "Authorization" = ("Basic {0}" -f $Auth);
        "Content-Type" = $_appType;
    };

    return $HeaderforAzure
}
  

function Azure_GetChildWorkItem
{
    param (
        [Int] $_WktID
        )

    $url = "https://analytics.dev.azure.com/$Azure_Organisation/$Azure_Project/_odata/v2.0//WorkItems?`$select=WorkItemId,Title,State,ChangedDate`&`$expand=Children(`$select=WorkItemId,Title,State,ChangedDate)&`$filter=WorkItemId%20eq%20$_WktID"

    $header = Azure_CreateHeader -_appType "application/json"
    $child = Invoke-RestMethod  -Method GET   -Uri $url  -Headers $header 
         
    return $child.value.children
}

function Azure_GetItem
{ param
    (
        [Int] $_ID
    )
        try
        {
            $url = "https://dev.azure.com/$Azure_Organisation/$Azure_Project/_apis/wit/workitems/$_ID`?api-version=5.1"
            $workitem = Invoke-RestMethod -Method GET -Uri $url -Headers $Azure_header        
            return $workitem
        }
        catch
        {
            Write-Log -Path $Config.logPath -Level "Error" -Message "$($MyInvocation.MyCommand) ::  Error while getting the existing workitems : $_"
        }
}

function SendMail
{
    param ($Body)
    $MailParam = @{
        To = $To
        From = $To
        Subject = "Morbois - Rapport Semaine $(get-date -Uformat '%V')"
        Body = $Body
        smtpserver = $SMTPServer
        Port = $SMTPPort
    }
    Send-MailMessage @MailParam -BodyAsHtml 
}
## Get Done Items from Azure


## Azure Variable
$Azure_Organisation = "****"
$Azure_Project = "****"
$Azure_Username = "****"
$Azure_Password = "****"
$Azure_Workitem_IDs = x,y,z

## Date Variables
$vendredi  = get-date 
$lundi = $vendredi.AddDays(-4)


#### Font Variables
$style_font = "font-family:Calibri"
$Font_Size =  "font-size:18px"

### Margin Variables
$Margin_Left_T1 = 15
$Margin_Left_T2 = $Margin_Left_T1 +15
$Margin_Left_T3 = $Margin_Left_T2 +15
$Margin_Top_Bottom = "margin-top:5;margin-bottom:5"


## Mail Variable
$To = "*****"
$SMTPServer =  "****"
$SMTPPort = 25

#### START 
$Email = "<div style =  $style_font;$Font_Size px> Hello, </br> `
Voici  le rapport de mon activit&eacute; de ma semaine du $(get-date $lundi -format 'dd') au $(get-date $vendredi  -format 'dd') $(get-date -Uformat '%B') $(get-date -format '%yy') </br>"

$Azure_header = Azure_CreateHeader -_appType "application/json-patch+json"

foreach($Epic_ID in $Azure_Workitem_IDs)
{
    $Epic = $false
    #Get Epic Workitem 
    $workitem = Azure_GetItem -_ID $Epic_ID

    #Get Epic Child -> User Story
    $US = Azure_GetChildWorkItem -_WktID $workitem.Id

    foreach ($SingleStory in $US)
    {
        if ($SingleStory.State -eq "Done")
        {
            $story = $false 

            $Tasks = Azure_GetChildWorkItem -_WktID $SingleStory.WorkItemId

            foreach ($Task in $Tasks)
            {
                # Avoid writting the Epic / story if nothing was done inside
                if (!($Epic))
                {
                    $Email += "<ul>`
                                <li  style = 'margin-left: $Margin_Left_T1`px;$Margin_Top_Bottom;$style_font;$Font_Size;list-style-type:disc;' > `
                                    <strong> $($workitem.fields.'System.Title') :</strong>`
                                </li>`
                            </ul>"
                    $epic = $true
                }
                if (!($story))
                {
                    $Email += "	<ul> <li style='margin-left: $Margin_Left_T2`px;$Margin_Top_Bottom;$style_font;$Font_Size;list-style-type:circle;'> <u>$($SingleStory.Title) :  </u></li> </ul>"
                    $story = $true
                }

                $limit = $(get-date ).addDays(-7)
                $changed = get-date  $($Task.ChangedDate) 
                if (($Task.State -eq "Done") -and ($changed -gt $limit))
                {
                    
                    $detail = Azure_GetItem -_ID $Task.WorkItemId
                    ## aligning the description on th esame line of the title
                    $description =  $($detail.fields.'System.Description').replace("<div>","").replace("</div>","")
                    $Email += " 	<ul> <li   style = 'margin-left: $Margin_Left_T3`px;$Margin_Top_Bottom;$style_font;$Font_Size;list-style-type:none;'> $($Task.Title) : $description </li></ul></br>"

                }
            }
        }
    }
}

$Email += "<div style =  $style_font;$Font_Size></br> Bon Week-end ! </br> Antoine</div>"

if ($DebugSoft.isPresent)
{
    $Email | out-file output.html 
}
else
{
    SendMail -Body $Email 
}
