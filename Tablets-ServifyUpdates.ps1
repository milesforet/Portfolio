[System.Collections.ArrayList]$errors = @()

try{
    #FRESHSERVICE API INFO
    $freshServiceBaseUrl = "https://abskids.freshservice.com"
    $freshApiKey = (Get-AutomationPSCredential -Name "service-fs-it@abskids.com").GetNetworkCredential().Password
    $freshApiKey = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($freshApiKey))

    #HEADER FOR FRESHSERVICE API CALLS
    $freshServiceHeaders = @{
        "Content-Type" = "application/json"
        "Authorization" = "Basic $freshApiKey"
    }

    $caseCustomObjectId = 21000042030
    [System.Collections.ArrayList]$tabletCases = @()

    $morePages = $true
    $url = "$freshServiceBaseUrl/api/v2/objects/$caseCustomObjectId/records?query=casestate%20%3A%20%27inProgress%27%20&per_page=100"

    Write-Output "`n`nGetting all custom object records."

    #GET ALL CUSTOM OBJECT RECORDS
    while($morePages){
        $customObjectResponse = Invoke-WebRequest -Method "Get" -Uri $url -Headers $freshServiceHeaders
        
        if($customObjectResponse.StatusCode -ne 200){
            throw "Failed to get custom object data... $($customObjectResponse.Content)"
        }

        $customObjectData = $customObjectResponse.Content | ConvertFrom-Json
        $tabletCases += $customObjectData.records
        

        if($customObjectResponse.RelationLink.next){
            $url = $customObjectResponse.RelationLink.next
            continue
        }

        $morePages = $false
    }

    $tabletCases = $tabletCases.Data

    Write-Output "`nNumber of cases with In Progress status: $($tabletCases.Count)`n"

    #CHECK IF THERE ARE ANY DUPLICATE SERIAL NUMBERS IN OPEN CASES
    $duplicates = $tabletCases.serialnumber | Group-Object | Where-Object {$_.Count -gt 1}
    if($duplicates){
        Write-Output "Number of duplicates: $($duplicates.Length)"
        foreach($dup in $duplicates.Name){
            Write-Output "`n===========$($dup)============`n"
            $duplicateCases = $tabletCases | Where-Object {$_.serialnumber -eq $dup}
            Write-Output "Cases: $($duplicateCases.reference_id -join ", ")"
            $duplicateCases | Select-Object -Property  "reference_id", "request_type", "casecreated", "status", "repairshipid", "repairshipstatus", "returnshipid", "returnshipstatus" | Format-Table -AutoSize
        
        $errors += "Duplicate case for $dup - cases: $($duplicateCases.reference_id -join ", ")"
        }
        Write-Output "Done with duplicates!`n"
    }

    $moreTicketPages = $true
    $ticketUrl = "$freshServiceBaseUrl/api/v2/tickets/filter?query=`"status:18`"&per_page=100"

    [System.Collections.ArrayList]$tickets = @()

    #GET ALL TICKETS 
    while($moreTicketPages){
        $ticketData = Invoke-WebRequest -Method "Get" -Uri $ticketUrl -Headers $freshServiceHeaders -SkipHttpErrorCheck

        if($ticketData.StatusCode -ne 200){
            throw "Failed to get tickets... $($ticketData.Content)"
        }

        $tickets += ($ticketData.Content | ConvertFrom-Json).tickets
        
        if($ticketData.RelationLink.next){
            $ticketUrl = $ticketData.RelationLink.next
            continue
        }

        $moreTicketPages = $false
    } 

    Write-Output "Number of tickets in Repair Status: $($tickets.Count)"
    Write-Output "`nGetting all custom object records and their associated tickets"

    if($tickets.Count -gt 100){
        $errors += "There are more than 100 tablet repair in progress tickets. The API wasn't returning next page link. Hopefully it is now!"
    }

    Write-Output "`n************************************************************************************************************"

    $ticketsWithServifyCases = @{}

    #LOOP THROUGH TICKETS
    foreach($ticket in $tickets){
        Write-Output "`n`n---------------$($ticket.id)--------------"
        

        if(!$ticket.custom_fields.msf_servify_cases){
            Write-Output "No custom object records selected."
        }else{
            Write-Output "This ticket has the following Custom Object IDs:"

            foreach($objRecordId in $ticket.custom_fields.msf_servify_cases){
                
                Write-Output "   #$objRecordId"

                if($ticketsWithServifyCases.ContainsKey($objRecordId)){
                    $ticketsWithServifyCases[$objRecordId] += $ticket.id
                    continue
                }

                $ticketsWithServifyCases[$objRecordId] = [System.Collections.ArrayList]@($ticket.id)
            }


        }

        #REGEX PATTERN FOR TABLET SERIAL NUMBERS
        $pattern = '\bR\d[A-Z0-9]{9,10}\b'

        #FIND SERIAL NUMBERS IN THE TICKETS
        if($ticket.custom_fields.tablet_serial_numbers){
            
            $serials = [regex]::Matches(($ticket.custom_fields.tablet_serial_numbers).ToUpper(), $pattern) | ForEach-Object { $_.Value } | Sort-Object -Unique

            if(!$serials){
                Write-Output "No Serial Numbers found in SN field"
                continue
            }

            [System.Collections.ArrayList]$newCasesOnTicket = @()

            #LOOP THROUGH THE SERIAL NUMBERS ASSOCIATED WITH THE TICKET AND CHECK IF THERE ARE ANY OPEN CASES THAT AREN'T TIED TO THE TICKET
            foreach($sn in $serials){

                Write-Output "`n$sn"

                #FIND OPEN CASES WITH MATCHING SN
                $casesForTicket = $tabletCases | Where-Object {$_.serialnumber -eq $sn -and $_.casestate -eq "inProgress"}

                if(!$casesForTicket){
                    Write-Output "No open case(s) found in custom object for this SN"
                    continue
                }

                Write-Output "Open case(s) found for this SN: $(($casesForTicket | Select-Object -Property bo_display_id, reference_id |
                        ForEach-Object { "$($_.bo_display_id) ($($_.reference_id))" }) -join ', ')"

                if($casesForTicket.Count -gt 1){
                    $errors += "$($sn) has multiple open cases: $($casesForTicket.bo_display_id -join ", ")"
                    continue
                }

                #IF OBJECT HAS CASE ALREADY ADDED, THAT MEANS IT'S ALREADY IN THE TICKET
                if($ticketsWithServifyCases.ContainsKey($casesForTicket.bo_display_id) -and $ticket.id -in $ticketsWithServifyCases[$casesForTicket.bo_display_id]){
                    Write-Output "$($casesForTicket.bo_display_id) ($($casesForTicket.reference_id)) is already logged."
                    continue
                }

                Write-Output "$($casesForTicket.bo_display_id) ($($casesForTicket.reference_id)) is not logged. That needs to be tied back to the ticket"

                #SINCE IT'S NOT TRACKED, THE CUSTOM OBJECT RECORD NEEDS TO BE ADDED TO THE TICKET
                if($ticketsWithServifyCases.ContainsKey($casesForTicket.bo_display_id)){
                    $ticketsWithServifyCases[$casesForTicket.bo_display_id] += $ticket.id
                }else{
                    $ticketsWithServifyCases[$casesForTicket.bo_display_id] = [System.Collections.ArrayList]@($ticket.id)
                }

                $newCasesOnTicket += $casesForTicket.bo_display_id
            }

            if($newCasesOnTicket){
                Write-Output "`nNeed to tie new case(s) to ticket"

                $updateCaseBody = @{
                    "custom_fields" = @{
                        "msf_servify_cases" = $newCasesOnTicket + $ticket.custom_fields.msf_servify_cases
                    }
                } | ConvertTo-Json

                $updateTicketRes = Invoke-WebRequest -Method Put -Uri "$freshServiceBaseUrl/api/v2/tickets/$($ticket.id)" -Headers $freshServiceHeaders -Body $updateCaseBody -SkipHttpErrorCheck
                
                if($updateTicketRes.StatusCode -ne 200){
                    Write-Output "FAILED UPDATING CUSTOM OBJECT FIELD"
                    Write-Output "BODY:`n$updateCaseBody"
                    Write-Output "`nRESPONSE:`n$($updateTicketRes.Content)"
                    $errors += "FAILED TO UPDATE $($ticket.id) with custom object data"
                    continue
                }

                Write-Output "Successfully updated servify custom object field!"
            }
        }
    }

    
    Write-Output "`n`n================  Ticket/CustObjRecords:  =================="
    $ticketsWithServifyCases
    Write-Output "`n============================================================"
    Write-Output "`nDone getting tickets with custom object records`n`n"
    Write-Output "************************************************************************************************************"


    $servifyCreds = Get-AutomationPSCredential -Name "Servify"
 
    #CREATE BODY AND HEADERS FOR SERVIFY REQUESTS
    $authBody = @{
        "LoginID" = $servifyCreds.UserName
        "Password" = $servifyCreds.GetNetworkCredential().Password
    } | ConvertTo-Json

    $servifyHeader = @{
        "accept" = "application/json"
        "accept-encoding" = "gzip, deflate, br, zstd"
        "accept-language" = "en-US,en;q=0.9"
        "app" = "SamsungB2B-Web"
        "appversion" = "1.5.1"
        "content-type" = "application/json"
        "languagecode" = "en"
        "origin" = "https://samsungenterprise.servify.tech"
        "priority" = "u=1, i"
        "referer" = "https://samsungenterprise.servify.tech/"
        "sec-ch-ua" = '"Not;A=Brand";v="99", "Microsoft Edge";v="139", "Chromium";v="139"'
        "sec-ch-ua-mobile" = "?0"
        "sec-ch-ua-platform" = "Windows"
        "sec-fetch-dest" = "empty"
        "sec-fetch-mode" = "cors"
        "sec-fetch-site" = "same-site"
        "timezone" = "-04:00"
        "user-agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36 Edg/139.0.0.0"
    }

    #GET SERVIFY TOKEN
    $authResponse = Invoke-WebRequest -Method POST -Uri "https://node-us.servify.tech/b2bGateway//api/v1/user/login" -Headers $servifyHeader -Body $authBody -SkipHttpErrorCheck

    if($authResponse.StatusCode -ne 200){
        throw "Auth response is not 200. $($authResponse.StatusCode) - $($authResponse.StatusDescription)"
    }

    $authContent = $authResponse | ConvertFrom-Json

    if($authContent.msg -ne "Success"){
        throw "Log in not successful `n`n $authContent"
    }


    #ADD TOKEN TO HEADER. WE'LL REUSE THE HEADER FOR OTHER REQUESTS, BUT NEED THE AUTH TOKEN
    $servifyHeader["authorization"] = $authContent.data.token
    
    #HASHTABLE OF HASHTABLES TO STORE TICKET UPDATE DATA
    $ticketUpdates = @{}

    foreach($custObjCase in $tabletCases){

        Write-Output "`n----------------$($custObjCase.reference_id) ($($custObjCase.serialnumber))----------------"


        #CREATE BODY WITH CASE NUMBER
        $servifyCaseBody = @{
            "ReferenceID" = $custObjCase.reference_id
        } | ConvertTo-Json

        #HTTP REQUEST FOR CASE DETAILS
        $caseDetails = Invoke-WebRequest -Method POST -Uri "https://node-us.servify.tech/b2bGateway//api/v1/csr/getDetails" -Headers $servifyHeader -Body $servifyCaseBody -SkipHttpErrorCheck

        $caseContent = $caseDetails.Content | ConvertFrom-Json

        if($caseDetails.StatusCode -ne 200 -or $caseContent.StatusCode -ne "SFY.B2B.2000"){
            Write-Output "$($custObjCase.reference_id) returned unsuccessful status code:`n`n$caseDetails"
            $errors.Add("$($custObjCase.reference_id) returned unsuccessful status code`n`n$caseDetails") | Out-Null
            continue
        }

        

        <# ---------------------------------EXAMPLE CUSTOM OBJECT RECORD FOR FIELD IDS
        casecreated      : 2026-03-16              -don't touch
        casestate        : closed                  -might need update
        reference_id     : RVCNSIIGLTBS            -don't touch
        repairshipid     :                         -only need to add on first update
        repairshipstatus :                         -continuous updates
        request_type     :                         -only need to add on first update
        returnshipid     :                         -continuous checks, 1 update
        returnshipstatus :                         -continuous updates
        serialnumber     : R9TX503HKSY             -don't touch
        status           : Service completed       -continuous updates
        ticket                                     -we'll see how to handle this one... probably an automator in fresh
        #>

        <#
            ADD CHECKS FOR THE FOLLOWING DATA... IF THERE IS A MISMATCH, WE'LL WANT TO UPDATE:
                -repairshipid
                -repairshipstatus
                -request_type
                -returnshipid
                -returnshipstatus
                -status
        #>

        

        $caseData = $caseContent.data.requestDetails

        #GET REQUEST TYPE (SERVIFY RETURNS DATA IN DIFFERENT IDS DEPENDING ON WHAT KIND OF CASE SO THIS TRANSLATES ALL THAT)
        $request_type = $caseData.RequestType.externalCode ?? $caseData.issue.IssueText
        $request_type = $request_type.Replace("_", " ")
        $request_type = $request_type -join ", "
        

        #CREATE HASH FOR CUSTOM OBJECT DATA AND CHECK WHAT NEEDS TO BE UPDATED
        $data = @{}

        [System.Collections.ArrayList]$caseUpdatesToTickets = @()

        #CHECK STATUS
        if($custObjCase.status -ne $caseData.currentStatus){
            Write-Output "Updating status to '$($caseData.currentStatus)'"
            $data["status"] = $caseData.currentStatus
            $caseUpdatesToTickets += "Status updated to '$($caseData.currentStatus)'"
        }

        #CHECK REPAIR REQUEST TYPE
        if($custObjCase.request_type -ne $request_type){
            Write-Output "Updating Request Type to $($request_type)"
            $data["request_type"] = $request_type
            $caseUpdatesToTickets += "Request Type updated from '$($custObjCase.request_type)' to '$request_type'"
        }

        #CHECK FOR BLANK/UPDATE TRACKING TO REPAIR FACILITY TRACKING NUMBER
        if($custObjCase.repairshipid -ne $caseData.partnerLogisticsWaybillnumber.customerToSCDetails.WayBillNumber){
            Write-Output "Updating RepairShipID to '$($caseData.partnerLogisticsWaybillnumber.customerToSCDetails.WayBillNumber)'"
            $data["repairshipid"] = $caseData.partnerLogisticsWaybillnumber.customerToSCDetails.WayBillNumber
            $caseUpdatesToTickets += "Shipment to Repair Center: $($caseData.partnerLogisticsWaybillnumber.customerToSCDetails.WayBillNumber)"
        }

        #CHECK REPAIR SHIPMENT STATUS
        if($custObjCase.repairshipstatus -ne $caseData.partnerLogisticsWaybillnumber.customerToSCDetails.Status){
            Write-Output "Updating RepairShipStatus to '$($caseData.partnerLogisticsWaybillnumber.customerToSCDetails.Status)'"
            $data["repairshipstatus"] = $caseData.partnerLogisticsWaybillnumber.customerToSCDetails.Status
            $caseUpdatesToTickets += "Shipment to Repair Center updated to '$($caseData.partnerLogisticsWaybillnumber.customerToSCDetails.Status)'"
        }

        #CHECK RETURN SHIPMENT ID
        if($custObjCase.returnshipid -ne $caseData.partnerLogisticsWaybillnumber.scToCustomerDetails.WayBillNumber){
            Write-Output "Updating ReturnShipID to '$($caseData.partnerLogisticsWaybillnumber.scToCustomerDetails.WayBillNumber)'"
            $data["returnshipid"] = $caseData.partnerLogisticsWaybillnumber.scToCustomerDetails.WayBillNumber
            $caseUpdatesToTickets += "Return shipment ID: $($caseData.partnerLogisticsWaybillnumber.scToCustomerDetails.WayBillNumber)"
        } 

        #CHECK RETURN STATUS
        if($custObjCase.returnshipstatus -ne $caseData.partnerLogisticsWaybillnumber.scToCustomerDetails.Status){
            Write-Output "Updating RepairShipStatus to '$($caseData.partnerLogisticsWaybillnumber.scToCustomerDetails.Status)'"
            $data["returnshipstatus"] = $caseData.partnerLogisticsWaybillnumber.scToCustomerDetails.Status
            $caseUpdatesToTickets += "Return shipment status updated to '$($caseData.partnerLogisticsWaybillnumber.scToCustomerDetails.Status)'"
        }

        #HANDLE CASE STATE
        $caseState = ""
        switch ($caseData.currentStatus) {
            {$_ -in @("Service completed", "Defective awaited")} {
                if($caseData.partnerLogisticsWaybillnumber.scToCustomerDetails.Status -in @("Device Delivered", "DELIVERED") ){
                    $caseState = "closed"
                }else{
                    $caseState = "inProgress"
                }
            }
            "Claim withdrawn" {$caseState = "cancelled"}
            "Service cancel" {$caseState = "cancelled"}
            "Device shipped" {$caseState = "inProgress"}
            "DropOff request accepted" {$caseState = "inProgress"}
            "Parts received" {$caseState = "inProgress"}
            "Repair completed" {$caseState = "inProgress"} 
            "Device received" {$caseState = "inProgress"}
            "Device dispatched" {$caseState = "inProgress"}
            "Replacement authorized" {$caseState = "inProgress"}
            Default {$custObjCase.currentStatus; Write-Output "$($caseData.currentStatus) is a new state. Please add to switch statement"; $errors.Add("New Case State ($($caseData.currentStatus)) found in $($custObjCase.reference_id)") | Out-Null}
        }

        if($custObjCase.casestate -ne $caseState -and $caseState -ne ""){
            Write-Output "Updating CaseState to $caseState"
            $data["casestate"] = $caseState
        }

        #SERVICE TYPE (THIS IS ALWAYS GOING TO BE MAIL IN UNLESS THEY REPLACE THE DEVICE)
        if($custObjCase.servicetype -ne $caseData.servicetype.Type){
            Write-Output "Updating ServiceType to $($caseData.servicetype.Type)"
            $data["servicetype"] = $caseData.servicetype.Type
        }

        #REPLACEMENT DEVICE
        if($custObjCase.replacementdevice -ne $caseContent.data.replacementDeviceDetails.productUniqueID){
            Write-Output "Updating ReplacementDevice to $($caseContent.data.replacementDeviceDetails.productUniqueID)"
            $data["replacementdevice"] = $caseContent.data.replacementDeviceDetails.productUniqueID
            $caseUpdatesToTickets += "Replacement device added: $($caseContent.data.replacementDeviceDetails.productUniqueID)"
        }
        

        if($data.Count -eq 0){
            Write-Output "Nothing to update!"
            continue
        }

        $updateRecord = @{
            "data" = $data
        } | ConvertTo-Json

        $updateResponse = Invoke-WebRequest -Method Put -Uri "$freshServiceBaseUrl/api/v2/objects/21000042030/records/$($custObjCase.bo_display_id)" -Headers $freshServiceHeaders -Body $updateRecord
        
        if($updateResponse.StatusCode -ne 200){
            Write-Output "ERROR UPDATING RECORD`n`nREQUEST BODY:`n$updateRecord`n`nRESPONSE:`n$($updateResponse.Content)"
            $errors.Add("Error updating case #$($custObjCase.reference_id)") | Out-Null
            continue
        }

        Write-Output "`nRECORD SUCCESSFULLY UPDATED!!"

        if($caseUpdatesToTickets -and $ticketsWithServifyCases[$custObjCase.bo_display_id]){
           
            $ticketIdsForUpdate = $ticketsWithServifyCases[$custObjCase.bo_display_id]
            Write-Output "`nThis case needs to update these tickets: $($ticketIdsForUpdate -join ", ")"

            foreach($tickId in $ticketIdsForUpdate){
                if($ticketUpdates.ContainsKey($tickId)){
                    $ticketUpdates[$tickId] += "<br><br><b>$($custObjCase.reference_id) - $($custObjCase.serialnumber)</b>:<br>$($caseUpdatesToTickets -join "<br>")"
                }else{
                    $ticketUpdates[$tickId] = "<b>$($custObjCase.reference_id) - $($custObjCase.serialnumber)</b>:<br>$($caseUpdatesToTickets -join "<br>")"
                }
            }
        }
    }

    Write-Output "`nDONE UPDATING CUSTOM OBJECT!"

    if($ticketUpdates){
        Write-Output "`n`nAdding notes to tickets..."

        foreach($tabletTicket in $ticketUpdates.Keys){

            Write-Output "`n------------  $tabletTicket  ------------"

            $addNoteBody = @{
                "private" = $true
                "body" = $ticketUpdates[$tabletTicket]
            } | ConvertTo-Json

            #CREATE PRIVATE NOTE ON TICKET
            $addNoteResponse = Invoke-WebRequest -Method Post -Uri "$freshServiceBaseUrl/api/v2/tickets/$tabletTicket/notes" -Body $addNoteBody -Headers $freshServiceHeaders -SkipHttpErrorCheck
            
            if($addNoteResponse.StatusCode -ne 201){
                Write-Output "Error creating private note - $($addNoteResponse.StatusDescription):"
                Write-Output $addNoteResponse.Content
                $errors += "Failed adding note to $tabletTicket"
                continue
            }

            Write-Output "SUCCESSFULLY ADDED NOTE!!"
        }
    }

}catch{
    $_
    $errors.Add($_) | Out-Null
}

try{
    #SEND EMAIL WITH ANY ERRORS
    if(!$errors){
        "Completed with no errors!!"
        exit
    }

    $smtpCreds = Get-AutomationPSCredential -Name "smtp2go"

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
        "Subject" = "Error(s) in Tablets-ServifyUpdates"
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