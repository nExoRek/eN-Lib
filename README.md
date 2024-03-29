**************************************************************
#  public script library https://w-files.pl

*a bunch of random scripts someone may find handful.*

## GENERAL 
  * [get-MACAddressVendor.ps1](https://github.com/nExoRek/eN-Lib/blob/master/get-MACAddressVendor.ps1) 
    connecting to macvendorloopkup and using IEEE OUI file to get NIC vendor
  
  * [get-MSeBooks.ps1](https://github.com/nExoRek/eN-Lib/blob/master/get-MSeBooks.ps1) 
    Microsoft eBook giveaway downloader. queries for available books and allows 
      you to download all or chosen titles

  * [show-myExternalIP.ps1](https://github.com/nExoRek/eN-Lib/blob/master/show-myExternalIP.ps1) 
    connects to whatsmyipaddress and provides information in the console

  * [compare-objectAttributeValues.ps1](https://github.com/nExoRek/eN-Lib/blob/master/compare-objectAttributeValues.ps1)
    compares two objects... couter-intuitively compare-object compares tables. this script is truely comparing objects. very useful!

  * [search-Windows.ps1](https://github.com/nExoRek/eN-Lib/blob/master/search-Windows.ps1)
    using Windows Search to locate file or folder.

  * [voice-Ping.ps1](https://github.com/nExoRek/eN-Lib/blob/master/voice-Ping.ps1)
    when you looking for some fun script to show PowerShell - this one can help. just don't foget to switch speakers on! (;

  * [MultiLevelTextMenu.lib.ps1](https://github.com/nExoRek/eN-Lib/blob/master/MultiLevelTextMenu.lib.ps1)
    library allowing to easily create text menu with multi-level choices (tree-menu)

## M365
  * [get-o365UserLicenseInformation.ps1](https://github.com/nExoRek/eN-Lib/blob/master/get-o365UserLicenseInformation.ps1)
    show information on user license focusing on direct/group assignment. particullarly
      usefull during moving from direct to Groub Based Licensing.
  * [get-ServicePlanInfo.ps1](https://github.com/nExoRek/eN-Lib/blob/master/get-ServicePlanInfo.ps1)
    handy tool for Service Plan lookups - SKU to friendly name resolution, which licenses contain given SP, list SPs included in license.
  * [get-AccountSkuReport.ps1](https://github.com/nExoRek/eN-Lib/blob/master/get-AccountSkuReport.ps1)
    create nice Account Sku report for documentation with display names.
  * [list-AADAdmins.ps1](https://github.com/nExoRek/eN-Lib/blob/master/list-AADAdmins.ps1)
    basic security report listing all priviledged accounts (but not PIM) and their statuses
  * [enable-MFAforUPN.ps1](https://github.com/nExoRek/eN-Lib/blob/master/enable-MFAforUPN.ps1)
    enforce MFA for a user. accepts pipelining of UPNs for bulk operation.
  * [set-mailboxOoOForwarding.ps1](https://github.com/nExoRek/eN-Lib/blob/master/set-mailboxOoOForwarding.ps1)
    another migration project support script. set up forward to target mailbox and OoO message. 
  * [disable-mailboxOoOForwarding.ps1](https://github.com/nExoRek/eN-Lib/blob/master/disable-mailboxOoOForwarding.ps1)
    reverse operation - disable forward rule and OoO.

## EXCHANGE
  * [get-SharedMailboxAccess.ps1](https://github.com/nExoRek/eN-Lib/blob/master/Get-SharedMailboxAccess.ps1) 
    prepares output for 'grant-SharedMailboxAccess' - use in source tenant, transform emails, use for target as import
  * [grant-SharedMailboxAccess.ps1](https://github.com/nExoRek/eN-Lib/blob/master/grant-SharedMailboxAccess.ps1) 
    suport script to grant "full access" + "send as" for a list of users and shared mailboxes. can as well be used as 
    import from source/backup.
  * [get-mobileDeviceReport.ps1](https://github.com/nExoRek/eN-Lib/blob/master/get-mobileDeviceReport.ps1)
    generate report on mobile devices registered under EXO users.

## AZURE
  * [search-AzureByIP.ps1](https://github.com/nExoRek/eN-Lib/blob/master/search-AzureByIP.ps1)
    searches Resources and Networks by IP
  
  * [search-AzureByName.ps1](https://github.com/nExoRek/eN-Lib/blob/master/search-AzureByName.ps1)
    searches Resources and resource-containers by name or partial name

  * [gui-ForceShutdownAzVM.ps1](https://github.com/nExoRek/eN-Lib/blob/master/gui-ForceShutdownAzVM.ps1)
    demo how to use elements of PS-GUI, force shutdown VM in Azure

  * [gui-ForceShutdownAzVMforms.ps1](https://github.com/nExoRek/eN-Lib/blob/master/gui-ForceShutdownAzVMforms.ps1)
    demo how to use element of Forms GUI in PS, force shutdown VM in Azure

  * [destroy-AzureVM.ps1](https://github.com/nExoRek/eN-Lib/blob/master/destroy-AzureVM.ps1)
    delete Azure VM along with related resources. v1 is limited - not giving an option to choose which resources. 

## AD
  * [remove-ProtectedOUStructure.ps1](https://github.com/nExoRek/eN-Lib/blob/master/remove-ProtectedOUStructure.ps1)
    remove whole OU structure removing 'Protect Object From Accidental Deletion' flag.
    
> nExoR 2o' ::))o-
  
**************************************************************
