Param($alertid,$subscriptionid)

# cleaning up the subscription name

$SubscriptionName = Get-SCOMNotificationSubscription -Id $subscriptionid | ft DisplayName | Out-String 
$SubscriptionName = $SubscriptionName.Substring($SubscriptionName.indexof(“-“)+12) 
$SubscriptionName = $SubscriptionName.trim() 
$SubscriptionName = “Processed by subscription: ” + $SubscriptionName

Get-SCOMAlert -Id ($alertid) | set-SCOMAlert –Comment $SubscriptionName -ResolutionState 100