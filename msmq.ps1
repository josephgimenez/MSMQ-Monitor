#Assign computer name to variable to avoid needing to change it in multiple locations later on
$computername = "xxxxxxxxxxx.domain.local"
#Assign counter variable to track number of error queue messages
$messagecount = 0

write-host "Importing pscx module for MSMQ commandlets...`n"

#import module that contains Get-MSMQueue commandlet
import-module pscx

#Call function and pass to foreach loop (%)
Get-MSMQueue | % { 

    #Initialize $msg
    $msg = ""
    #Append name of queue to computer name
    $queuefull = $computername + "\" + $_.QueueName
    
    #Create new MessageQueue object based on current queue in loop
    $QueueMsg = new-object System.Messaging.MessageQueue $queuefull

    #If total messages in queue > 0 and "error" found in name of queue
    if ($_.QueueName.contains("error"))
    { 
	#write-host "Queue Name: " + $_.QueueName

       	if ($QueueMsg.GetAllMessages().Length -gt 0)
	{ 
		#Assign log entry string (date + queue name + queue message count)
		$msg =  $(get-date).ToString() + " --> " + $_.QueueName + " has " + $QueueMsg.GetAllMessages().Length + " messages.`n"
		#Output message to standard output (console)
		write-host $msg
		#Append log entry to queue log file
		add-content c:\scripts\messagequeuelog.txt "----------"
		add-content c:\scripts\messagequeuelog.txt $msg
		#Increment total message count for error queues
		$messagecount += $QueueMsg.GetAllMessages().Length

		$QueueMsg.GetAllMessages() | % {
			$msgID = "Message ID: " + $_.Id;
			add-content c:\scripts\messagequeuelog.txt $msgID;
		}


		#format msg to work properly with eventcreate command
		#$msg = "'<BR>" + $msg + "<BR>Ran cmd 'ReturnToSourceQueue' to return message to normal queue.'"
		#

		# is it the tax credit queue?
		if ($_.QueueName.contains("taxcredit"))
		{
			# Flush error queue messages back to normal queue
			#C:\nservicebus.2.0.0.1219\tools\ReturnToSourceQueue.exe taxcrediterrorqueueproduction all	
			#Event Log entry for 9111 - Tax Credit Queue
			eventcreate /ID 911 /L Application /T Information /SO MSMQTaxCreditError /D $msg 
			#Log Entry to notify user we flushed the queue
			#add-content c:\scripts\messagequeuelog.txt "Ran cmd 'ReturnToSourceQueue' to return message to normal queue.`n"
		}

		if ($_.QueueName.contains("assessment"))
		{
			#C:\nservicebus.2.0.0.1219\tools\ReturnToSourceQueue.exe assessmenterrorqueueproduction all
			eventcreate /ID 912 /L Application /T Information /SO MSMQAssessmentError /D $msg 
			#add-content c:\scripts\messagequeuelog.txt "Ran cmd 'ReturnToSourceQueue' to return message to normal queue.`n"
		}

		if ($_.QueueName.contains("communications"))
		{
			#C:\nservicebus.2.0.0.1219\tools\ReturnToSourceQueue.exe communicationserrorqueueproduction all
			eventcreate /ID 913 /L Application /T Information /SO MSMQCommsError /D $msg 
			#add-content c:\scripts\messagequeuelog.txt "Ran cmd 'ReturnToSourceQueue' to return message to normal queue.`n"
		}

		if ($_.QueueName.contains("shrawe"))
		{
			#C:\nservicebus.2.0.0.1219\tools\ReturnToSourceQueue.exe shraweberrorqueueproduction all
			eventcreate /ID 914 /L Application /T Information /SO MSMQShraweError /D $msg 
			#add-content c:\scripts\messagequeuelog.txt "Ran cmd 'ReturnToSourceQueue' to return message to normal queue.`n"
		}

		if ($_.QueueName.contains("integration"))
		{
			eventcreate /ID 915 /L Application /T Information /SO MSMQIntegError /D $msg
		}

		if ($_.QueueName.contains("document"))
		{
			eventcreate /ID 916 /L Application /T Information /SO MSMQDocError /D $msg
		}

		if ($_.QueueName.contains("mediaprocess"))
		{
			eventcreate /ID 917 /L Application /T Information /SO MSMQMediaError /D $msg
		}

		if ($_.QueueName.contains("people"))
		{
			eventcreate /ID 918 /L Application /T Information /SO MSMQPeopleError /D $msg
		}

		if ($_.QueueName.contains("verify"))
		{
			eventcreate /ID 919 /L Application /T Information /SO MSMQEVerifyError /D $msg
		}

		if ($_.QueueName.contains("scheduler"))
		{
			eventcreate /ID 920 /L Application /T Information /SO MSMQSchedulerError /D $msg
		}

		add-content c:\scripts\messagequeuelog.txt "----------`n"
	}
    }
}

#Are there exactly zero total messages in all the queues?
if ($messagecount -eq 0)
{
    write-host "There are no messages in the error queues!"
}
