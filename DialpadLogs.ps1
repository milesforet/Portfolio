[System.Collections.ArrayList]$errors = @()


try{

    Import-Module "PnP.Powershell"

    #FRESHSERVICE API INFO
    $freshServiceBaseUrl = "https://abskids.freshservice.com"
    $freshApiKey = (Get-AbsCred -credName "service-fs-it@abskids.com").GetNetworkCredential().Password
    $freshApiKey = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($freshApiKey))

    #HEADER FOR FRESHSERVICE API CALLS
    $freshServiceHeaders = @{
        "Content-Type" = "application/json"
        "Authorization" = "Basic $freshApiKey"
    }


    [System.Collections.ArrayList] $offboardTickets = @()

    #CHECK FOR TAG AND STATUS IS NOT RESOLVED
    $offboardUrl = "$freshServiceBaseUrl/api/v2/tickets/filter?workspace_id=12&query=`"tag:%27VOIP-Offboard%27%20AND%20tag:%27need_voip_logs%27%20AND%20status:<3`"&per_page=100"

    $morePages = $true
    while($morePages){
        $dialpadOffboardTickets = Invoke-WebRequest -Method "Get" -Uri $offboardUrl -Headers $freshServiceHeaders

        if($dialpadOffboardTickets.RelationLink.next){
            Write-Output "There are more pages!"
            $offboardUrl = $dialpadOffboardTickets.RelationLink.next
        }else{
            Write-Output "No more pages"
            $morePages = $false
        }

        $offboardTickets += ($dialpadOffboardTickets.Content | ConvertFrom-Json).tickets
        
        if([int32]($dialpadOffboardTickets.Headers."x-ratelimit-remaining")[0] -lt 5){
                Write-Output "Pausing for Fresh ratelimit"
                Start-Sleep -Seconds 60
        }
    }


    
    Write-Output "Total number of tickets that need to be processed: $($offboardTickets.Count)"

    if($offboardTickets.Count -eq 0){
        Write-Output "Exiting since there are no tickets to work."
        break
    }
    
    # WE ARE ONLY DOING 4 AT A TIME TO NOT HIT 200/HR POST RATE LIMIT ON POSTING TO STATS ENDPOINT
    if($offboardTickets.Count -gt 5){
        $errors += "More than 5 tickets, we are processing the first 5: $($offboardTickets.id -join ", ")"

        $offboardTickets = @($offboardTickets | Sort-Object -Property "due_by" | Select-Object -Property id, due_by -First 5)
    }


    Write-Output "Number of tickets we're actually processing: $($offboardTickets.Count)"

    $dialpadApiKey = (Get-AbsCred -credName "dialpad-admin").GetNetworkCredential().Password
    $dialpadHeader = @{
        "accept" = "application/json"
        "authorization" = "Bearer $dialpadApiKey"
    }

    
    $spCreds = Get-AbsCred -credName "service-provisioning"

    try{
        $conn = Get-PnPConnection
        Write-Output "Already connected to SP"
    }catch{
        Write-Output "NOT CONNECTED TO SP... Connecting"
        Connect-PnPOnline -Url "https://abs79.sharepoint.com/sites/VOIPLogs" -Credentials $spCreds -ClientId "e41d925a-fe12-4c7f-9675-87a1e5a04e7d"
        Write-Output "Successfully connected to SP`n"
    }


    [hashtable]$offboardData = @{}

    # LOOP THROUGH TICKETS, GET SERVICE ITEM INFO, TRIGGER DIALPAD LOG EXPORT
    # IF SUCCESSFUL, ADD TAG AND ADD LINKS FOR GETTING DIALPAD OFFBOARD LOGS
    foreach($ticket in $offboardTickets){

        try{
            Write-Output "`n############# $($ticket.id) #############"

            # GET INFO FROM SERVICE REQUEST ITEM
            $requestInfoResponse = Invoke-WebRequest -Method "Get" -Uri "$freshServiceBaseUrl/api/v2/tickets/$($ticket.id)/requested_items" -Headers $freshServiceHeaders

            if([int32]($requestInfoResponse.Headers."x-ratelimit-remaining")[0] -lt 5){
                Start-Sleep -Seconds 60
                Write-Output "Pausing for Fresh ratelimit"
            }

            $requestInfo = ($requestInfoResponse | ConvertFrom-Json).requested_items
            
            $name = $requestInfo.custom_fields.employee_name_on_account
            $email = $requestInfo.custom_fields.email_address

            Write-Output "$name - $email`n"

            # GET TAGS SINCE FRESH DOESNT LET ME GET IT FROM FILTER AND DOES HAVE AN OPTION TO APPEND A TAG. IT OVERWRITES THE CURRENT TAGS SO EXISTING NEED TO BE INCLUDED
            $tagsResponse = Invoke-WebRequest -Method "Get" -Uri "$freshServiceBaseUrl/api/v2/tickets/$($ticket.id)?include=tags" -Headers $freshServiceHeaders
            $ticketTags = ($tagsResponse.Content | ConvertFrom-Json).ticket.tags

            #GET USER FROM DIALPAD
            $dialpadUserResponse = Invoke-WebRequest -Method "Get" -Uri "https://dialpad.com/api/v2/users?email=$email" -Headers $dialpadHeader
            $dialpadUser = ($dialpadUserResponse | ConvertFrom-Json).items

            Write-Output "Dialpad User ID: $($dialpadUser.id)"

            Write-Output "`n"

            # IF NO DIALPAD ACCOUNT
            if(!$dialpadUser.id){
                Write-Output "NOT FOUND IN DIALPAD"

                $newTags = $ticketTags -ne "need_voip_logs"

                $addTagBody = @{
                        "tags" = $newTags
                } | ConvertTo-Json

                $addTagResponse = Invoke-WebRequest -Method "Put" -Uri "$freshServiceBaseUrl/api/v2/tickets/$($ticket.id)" -Headers $freshServiceHeaders -Body $addTagBody
                Write-Output "Removed tag"

                $noteBody = @{
                    "private" = $true
                    "body" = "User not found in Dialpad"
                } | ConvertTo-Json

                $noteResponse = Invoke-WebRequest -Method POST "$freshServiceBaseUrl/api/v2/tickets/$($ticket.id)/notes" -Headers $freshServiceHeaders -Body $noteBody
                Write-Output "Private note added"

                continue
            }

            # DAYS OF HOW FAR WE NEED TO LOOK BACK (AKA THE DAY THEIR DP ACCOUNT WAS CREATED)
            $addedToDpDaysAgo = ((Get-Date)-$dialpadUser.date_added).Days

            $offboardDate = [datetime]$requestInfo.custom_fields.end_date

            # DAYS OF WHEN WE WANT TO START THE QUERY (AKA THE OFFBOARD DATE)
            $offboardDaysAgo = ((Get-Date)-$offboardDate).days

            Write-Output "Offboard Date: $($offboardDate.ToString("MM/dd/yyyy"))"
            Write-Output "DP Account created: $(($dialpadUser.date_added).ToString("MM/dd/yyyy"))`n"

            Write-Output "# of days ago we're starting the query: $offboardDaysAgo"
            Write-Output "# of days we're looking back: $addedToDpDaysAgo"
        

            [System.Collections.ArrayList]$statBodies = @()    


            # IF TOO LONG, WE NEED TO BREAK UP INTO SMALLER DATE RANGES
            if(($addedToDpDaysAgo-$offboardDaysAgo) -gt 75){
                Write-Output "Log length would be more than 75 days. Breaking into chunks"
                Write-Output "`nLog chunks:"

                $logStartDate = $offboardDaysAgo
                $more = $true

                while($more){
                    
                    $logEndDate = $logStartDate+75

                    # IF LOG DATE GOES PAST WHEN DP ACCOUNT WAS CREATED, WE PULL IT BACK TO CREATED DATE
                    if($logEndDate -gt $addedToDpDaysAgo){
                        $logEndDate = $addedToDpDaysAgo
                        $more = $false
                    }

                    Write-Output "$logStartDate-$logEndDate"

                    $statBodies += @{
                        "export_type" = "records"
                        "days_ago_start" = $logStartDate
                        "days_ago_end" = $logEndDate
                        "target_type" = "user"
                        "target_id" = $dialpadUser.id
                        "timezone" = "UTC" ################### WE NEED TO LOOK AT THIS??????? POTENTIALLY USE THE USERS TIMEZONE?
                    }

                    $logStartDate += 76
                }
            }else{
                Write-Output "Log length is less than 75 days. No chunks necessary"

                $statBodies += @{
                    "export_type" = "records"
                    "days_ago_start" = $offboardDaysAgo
                    "days_ago_end" = $addedToDpDaysAgo
                    "target_type" = "user"
                    "target_id" = $dialpadUser.id
                    "timezone" = "UTC"
                }
            }

            # ALL THE STATS WE'RE PULLING
            $statsToPull = @("calls", "recordings", "texts", "voicemails")
            $dialStatsUrl = "https://dialpad.com/api/v2/stats"

            $allStats = @{}

            # START STAT EXPORT FOR EACH TYPE
            foreach($stat in $statsToPull){
                Write-Output "`n--- $stat ---"
                $allStats[$stat] = [System.Collections.ArrayList]@()
                foreach($statBody in $statBodies){
                    try{
                    $statBody["stat_type"] = $stat

                    #$statBody | ConvertTo-Json

                    $statResponse = Invoke-WebRequest -Method "Post" -Uri $dialStatsUrl -Headers $dialpadHeader -Body $statBody
                    #Write-Output "API Call - $((Get-Date).ToString("MM-dd-yyyy hh:mm tt"))"
                    $statId = ($statResponse.Content | ConvertFrom-Json).request_id

                    Write-Output "Success! ID: $statId"
                    
                    $allStats[$stat] += $statId

                    }catch{
                        $_
                        $errors += "$($ticket.id) ($name - $email) :`n$_"
                    }
                }
                
            }

            $offboardData[[string]$ticket.id] = @{
                "UploadFolder" = ($requestInfo.custom_fields.voip_log_link -split "Offboard%20Logs/")[1]
                "Logs" = $allStats
                "User" = "$name - $email"
                "Tags" = $ticketTags
                "OffboardDate" = ($requestInfo.custom_fields.end_date).Substring(5)
                "Username" = $email.Replace("@abskids.com", "")
            }

            Write-Output ""

        }catch{
            $_
            $errors += $_
        }
    }

    Write-Output "================================================ NOW WE GET THE LOGS AND UPLOAD THEM ================================================"

    # NEED TO LOOP THROUGH THE HASH TO CHECK IF ALL THE DATA IS AVAILABLE
    foreach($ticket in $offboardData.Keys){
        try{
            Write-Output "`n`n############# $ticket #############"
            Write-Output $offboardData[$ticket]["User"]

            # LOOP THROUGH OUR HASH THAT HAS ALL THE DATA ON OFFBOARD TICKETS WE NEED
            foreach($statType in $offboardData[$ticket]["Logs"].Keys){
                Write-Output "`n----- $statType -----"

                [System.Collections.ArrayList]$currentFile = @()
                $capturedHeader = $null

                foreach($record in $offboardData[$ticket]["Logs"][$statType]){
                    Write-Output "URL: $("https://dialpad.com/api/v2/stats/$record")"

                    $waiting = $true

                    $logsData = $null

                    # LOOP WHILE WE WAIT FOR EXPORT TO BE AVAILABLE FOR DOWNLOAD
                    while($waiting){
                        
                        $logsResponse = Invoke-WebRequest -Uri "https://dialpad.com/api/v2/stats/$record" -Headers $dialpadHeader
                        $logsData = ($logsResponse.Content | ConvertFrom-Json)

                        # IF LOGS AREN'T AVAILABLE, SLEEP FOR 15 SECONDS
                        if($logsData.status -eq "processing"){
                            Write-Output "Current Status: $($logsData.status) | Sleeping for 15 Seconds | $((Get-Date).ToSTring("MM-dd-yyyy hh:mm tt"))"
                            Start-Sleep -Seconds 15
                            continue
                        }elseif ($logsData.status -eq "complete") {

                            $waiting = $false
                            "Current Status: $($logsData.status)"

                            $csvLogs = Invoke-WebRequest -Uri $logsData.download_url -Headers $dialpadHeader

                            $memoryStream = New-Object System.IO.MemoryStream(,$csvLogs.Content)
                            $reader = [System.IO.StreamReader]::new($memoryStream)
                            $rawCsvText = $reader.ReadToEnd()

                            if (-not $capturedHeader -and -not [string]::IsNullOrWhiteSpace($rawCsvText)) {
                                $capturedHeader = ($rawCsvText -split "`r?`n")[0]
                            }

                            $dataAsCSV = $rawCsvText | ConvertFrom-Csv
                            if($dataAsCSV){
                                if ($dataAsCSV -is [array]) {
                                    $currentFile.AddRange($dataAsCSV)
                                } else {
                                    $currentFile.Add($dataAsCSV) | Out-Null
                                }
                            }
                            
                            Write-Output "Row count: $($currentFile.Count)"

                        }else{
                            Write-Output "$ticket - Unexpected status: $($logsData.status)"
                            throw "$ticket - Unexpected status: $($logsData.status)"
                        }

                        

                    }

                }
                if($currentFile.Count -gt 1){
                    switch($statType){
                        "calls" {$currentFile = $currentFile | Sort-Object -Property "date_started" -Descending}
                        "recordings" {$currentFile = $currentFile | Sort-Object -Property "date" -Descending}
                        "texts" {$currentFile = $currentFile | Sort-Object -Property "date" -Descending}
                        "voicemails" {$currentFile = $currentFile | Sort-Object -Property "date" -Descending}
                    }
                }
                

                # ADD FILE TO SHAREPOINT

                $uploadFolder = $offboardData[$ticket]["UploadFolder"]

                if ($currentFile.Count -gt 0) {
                    $csvRawText = $currentFile | ConvertTo-Csv -NoTypeInformation | Out-String
                }
                elseif ($capturedHeader) {
                    $csvRawText = $capturedHeader + "`r`n"
                }
                else {
                    $csvRawText = $null
                }

                $bytes = [System.Text.Encoding]::UTF8.GetBytes($csvRawText)
                $uploadStream = New-Object System.IO.MemoryStream(,$bytes)

                $i = Add-PnpFile -Stream $uploadStream -Folder "Shared Documents/Offboard Logs/$uploadFolder" -FileName "$ticket-$($offboardData[$ticket]["Username"])-$statType.csv"
                Write-Output "File Uploaded!"


            }

            Write-Output "`n"

            $newTags = ($offboardData[$ticket]["Tags"] + "voip_logs_done") -ne "need_voip_logs"

            $addTagBody = @{
                    "tags" = $newTags
            } | ConvertTo-Json

            $addTagResponse = Invoke-WebRequest -Method "Put" -Uri "$freshServiceBaseUrl/api/v2/tickets/$($ticket)" -Headers $freshServiceHeaders -Body $addTagBody
            Write-Output "Tag added to mark logs done"

            $noteBody = @{
                "private" = $true
                "body" = "Logs exported and uploaded to Sharepoint"
            } | ConvertTo-Json

            $noteResponse = Invoke-WebRequest -Method POST "$freshServiceBaseUrl/api/v2/tickets/$ticket/notes" -Headers $freshServiceHeaders -Body $noteBody
            Write-Output "Private note added"

        }catch{
            $_
            $errors += $_
        }
    }

}catch{
    $_
    $errors += $_
} 
Write-Output "`n"

exit

#HANDLE ERRORS
try{
    #SEND EMAIL WITH ANY ERRORS
    if(!$errors){
        "Completed with no errors!!"
        exit
    }

    $smtpCreds = Get-AbsCred -credName "smtp2go"

    $errorEmailBody = ""

    foreach($err in $errors){
        $errorEmailBody += "<p>$err</p><br>"   
    }

    $emailHeaders = @{
        "Content-Type" = "application/json"
        "X-Smtp2go-Api-Key" = $smtpCreds.GetNetworkCredential().Password
        "Accept" = "application/json"
    }
    
    $emailBody = @{
        "sender" = "mforet@abskids.com"
        "to" = @("mforet@abskids.com")
        "Subject" = "Error(s) in FS-DialpadLogs"
        "html_body" = $errorEmailBody
    } | ConvertTo-Json

    $sendEmailRes = Invoke-WebRequest -Method "POST" -Uri "https://api.smtp2go.com/v3/email/send" -Headers $emailHeaders -Body $emailBody -SkipHttpErrorCheck


    if($sendEmailRes.StatusCode -ne 200){
        throw "Failed to send errror email"
    }

    Write-Output "Email with errors sent!"

}catch{
    $_
    exit
}   
