<#
.SYNOPSIS
    returns information on user license type and assignment source
.DESCRIPTION
    scirpt simplifies verification of license assignments. especially usefull during
    moving to Group Based Licensing. it allows to be run for a single user by UPN name
    or directly from get-MSOLuser output to generate report
    based on examples from
    https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-ps-examples
.EXAMPLE
    .\get-o365UserLicenseInformation.ps1 -UserPrincipalName nexor@w-files.pl
    
    get license information on a single user
.EXAMPLE
    cat c:\temp\userList.txt|.\get-o365UserLicenseInformation.ps1 -includeServicePlans -exportToCSV
    
    get license information on all users in flat, txt file containing UPNs. then exports to csv with
    default name of 'LicenseReport-<date>.csv' comma-delimited.
    results will contain detailed service plan information. service plans are array object.
.EXAMPLE
    cat c:\temp\userList.txt|.\get-o365UserLicenseInformation.ps1|export-csv -nti myReportName.csv -delimiter ';'
    
    get license information on all users in flat, txt file containing UPNs. then exports to csv with
    custom options - delimiter and file name.
.EXAMPLE
    Get-MsolUser -SearchString ne|.\get-o365UserLicenseInformation.ps1|export-csv -nti -deli ';' ne-users.csv
    
    script accepts msol user objects. in that example it will look for users that emails or UPNs 
    begins with 'ne' and exports data to csv, semicolon delimited. 
.EXAMPLE
    Get-MsolUser -all|.\get-o365UserLicenseInformation.ps1 -includeServicePlans -exportToCSV
    
    script accepts msol user objects. in that example it will export all user licensing information.
    such report may help you during switchover to GBL.
.OUTPUTS
    PSCustomObject @{
                userPrincipalName  
                AccountSkuId
                assignSource
                licenseName 
                licenseGroup
                usageLocation
                [servicePlans]
            }
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201015
    - 202015 usageLocation, -includeServicePlans, connection check fix, licenseName fix, -exportToCSV
    - 201012 v1

    TO|DO
    - change licsku and licgroups to hashtables
#>
#Requires -Module MSOnline
[cmdletbinding(DefaultParameterSetName="UPN")]
param(
    #get information on user with given UPN
    [parameter(mandatory=$true,position=0,ValueFromPipeline=$true,ParameterSetName="UPN")]
        [alias("UPN")]
        [string]$UserPrincipalName,
    #gets information directly from MSOLUser object - especially useful during pipelining
    [parameter(mandatory=$true,position=0,ValueFromPipeline=$true,ParameterSetName="MSOL")]
        [Microsoft.Online.Administration.User]$MSOLUser,
    #extended report including service plans information
    [Parameter(mandatory=$false,position=1)]
        [switch]$includeServicePlans,
    #automatically exports to CSV using some predefined values. use this option, or pipe on 
    #export-csv with your own parameters
    [Parameter(mandatory=$false,position=2)]
        [switch]$exportToCSV
)

begin {
    function get-LicenseSKUs {
        param()
        Write-Verbose "getting tenant SKUs..."
        $LicenseSKUs=@()
        try{
            Get-MsolAccountSku -ErrorAction stop|%{
                $LicenseSKUs+= [PSCustomObject]@{
                    SKUID=$_.SkuId
                    AccountSkuId=$_.AccountSkuId
                    SKUPartNumber=$_.SkuPartNumber
                    AccountObjectId=$_.AccountObjectId
                } 
            }
            return $LicenseSKUs
        } catch {
            write-host -ForegroundColor Red $_.Exception
            exit -1
        }
    }
    function get-LicenseGroups {
        param()
        Write-Verbose "getting Licensing Groups..."
        $licGroups= Get-MsolGroup -All | Where-Object {$_.Licenses} | Select-Object ObjectId, DisplayName, `
            @{Name="Licenses";Expression={$_.Licenses | Select-Object -ExpandProperty SkuPartNumber}}
        if( [string]::IsNullOrEmpty($licGroups) ) { 
            Write-Warning "no GBL groups found...."
            return $null
        }
        return $licGroups
    }
    #https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-service-plan-reference
    #region SKUNAMES
    $listOfAllLicenseSKUNames=@() 
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="SPZA_IW";Name="APP CONNECT IW";GUID="8f0c5670-4e56-4892-b06d-91c085d7004f"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MCOMEETADV";Name="AUDIO CONFERENCING";GUID="0c266dff-15dd-4b49-8397-2bb16070ed52"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="AAD_BASIC";Name="AZURE ACTIVE DIRECTORY BASIC";GUID="2b9c8e7c-319c-43a2-a2a0-48c5c6161de7"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="AAD_PREMIUM";Name="AZURE ACTIVE DIRECTORY PREMIUM P1";GUID="078d2b04-f1bd-4111-bbd4-b4b1b354cef4"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="AAD_PREMIUM_P2";Name="AZURE ACTIVE DIRECTORY PREMIUM P2";GUID="84a661c4-e949-4bd2-a560-ed7766fcaf2b"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="RIGHTSMANAGEMENT";Name="AZURE INFORMATION PROTECTION PLAN 1";GUID="c52ea49f-fe5d-4e95-93ba-1de91d380f89"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="DYN365_ENTERPRISE_PLAN1";Name="DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN ENTERPRISE EDITION";GUID="ea126fc5-a19e-42e2-a731-da9d437bffcf"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="DYN365_ENTERPRISE_CUSTOMER_SERVICE";Name="DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION";GUID="749742bf-0d37-4158-a120-33567104deeb"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="DYN365_FINANCIALS_BUSINESS_SKU";Name="DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION";GUID="cc13a803-544e-4464-b4e4-6d6169a138fa"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE";Name="DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION";GUID="8edc2cf8-6438-4fa9-b6e3-aa1660c640cc"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="DYN365_ENTERPRISE_SALES";Name="DYNAMICS 365 FOR SALES ENTERPRISE EDITION";GUID="1e1a282c-9c54-43a2-9310-98ef728faace"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="DYN365_ENTERPRISE_TEAM_MEMBERS";Name="DYNAMICS 365 FOR TEAM MEMBERS ENTERPRISE EDITION";GUID="8e7a3d30-d97d-43ab-837c-d7701cef83dc"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="Dynamics_365_for_Operations";Name="DYNAMICS 365 UNF OPS PLAN ENT EDITION";GUID="ccba3cfe-71ef-423a-bd87-b6df3dce59a9"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="EMS";Name="ENTERPRISE MOBILITY + SECURITY E3";GUID="efccb6f7-5641-4e0e-bd10-b4976e1bf68e"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="EMSPREMIUM";Name="ENTERPRISE MOBILITY + SECURITY E5";GUID="b05e124f-c7cc-45a0-a6aa-8cf78c946968"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="EXCHANGESTANDARD";Name="EXCHANGE ONLINE (PLAN 1)";GUID="4b9405b0-7788-4568-add1-99614e613b69"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="EXCHANGEENTERPRISE";Name="EXCHANGE ONLINE (PLAN 2)";GUID="19ec0d23-8335-4cbd-94ac-6050e30712fa"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="EXCHANGEARCHIVE_ADDON";Name="EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE";GUID="ee02fd1b-340e-4a4b-b355-4a514e4c8943"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="EXCHANGEARCHIVE";Name="EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER";GUID="90b5e015-709a-4b8b-b08e-3200f994494c"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="EXCHANGEESSENTIALS";Name="EXCHANGE ONLINE ESSENTIALS";GUID="7fc0182e-d107-4556-8329-7caaa511197b"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="EXCHANGE_S_ESSENTIALS";Name="EXCHANGE ONLINE ESSENTIALS";GUID="e8f81a67-bd96-4074-b108-cf193eb9433b"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="EXCHANGEDESKLESS";Name="EXCHANGE ONLINE KIOSK";GUID="80b2d799-d2ba-4d2a-8842-fb0d0f3a4b82"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="EXCHANGETELCO";Name="EXCHANGE ONLINE POP";GUID="cb0a98a8-11bc-494c-83d9-c1b1ac65327e"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="INTUNE_A";Name="INTUNE";GUID="061f9ace-7d42-4136-88ac-31dc755f143f"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="M365EDU_A1";Name="Microsoft 365 A1";GUID="b17653a4-2443-4e8c-a550-18249dda78bb"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="M365EDU_A3_FACULTY";Name="Microsoft 365 A3 for faculty";GUID="4b590615-0888-425a-a965-b3bf7789848d"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="M365EDU_A3_STUDENT";Name="Microsoft 365 A3 for students";GUID="7cfd9a2b-e110-4c39-bf20-c6a3f36a3121"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="M365EDU_A5_FACULTY";Name="Microsoft 365 A5 for faculty";GUID="e97c048c-37a4-45fb-ab50-922fbf07a370"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="M365EDU_A5_STUDENT";Name="Microsoft 365 A5 for students";GUID="46c119d4-0379-4a9d-85e4-97c66d3f909e"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="O365_BUSINESS";Name="MICROSOFT 365 APPS FOR BUSINESS";GUID="cdd28e44-67e3-425e-be4c-737fab2899d3"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="SMB_BUSINESS";Name="MICROSOFT 365 APPS FOR BUSINESS";GUID="b214fe43-f5a3-4703-beeb-fa97188220fc"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="OFFICESUBSCRIPTION";Name="MICROSOFT 365 APPS FOR ENTERPRISE";GUID="c2273bd0-dff7-4215-9ef5-2c7bcfb06425"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="O365_BUSINESS_ESSENTIALS";Name="MICROSOFT 365 BUSINESS BASIC";GUID="3b555118-da6a-4418-894f-7df1e2096870"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="SMB_BUSINESS_ESSENTIALS";Name="MICROSOFT 365 BUSINESS BASIC";GUID="dab7782a-93b1-4074-8bb1-0e61318bea0b"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="O365_BUSINESS_PREMIUM";Name="MICROSOFT 365 BUSINESS STANDARD";GUID="f245ecc8-75af-4f8e-b61f-27d8114de5f3"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="SMB_BUSINESS_PREMIUM";Name="MICROSOFT 365 BUSINESS STANDARD";GUID="ac5cef5d-921b-4f97-9ef3-c99076e5470f"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="SPB";Name="MICROSOFT 365 BUSINESS PREMIUM";GUID="cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="SPE_E3";Name="MICROSOFT 365 E3";GUID="05e9a617-0261-4cee-bb44-138d3ef5d965"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="SPE_E5";Name="Microsoft 365 E5";GUID="06ebc4ee-1bb5-47dd-8120-11324bc54e06"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="SPE_E3_USGOV_DOD";Name="Microsoft 365 E3_USGOV_DOD";GUID="d61d61cc-f992-433f-a577-5bd016037eeb"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="SPE_E3_USGOV_GCCHIGH";Name="Microsoft 365 E3_USGOV_GCCHIGH";GUID="ca9d1dd9-dfe9-4fef-b97c-9bc1ea3c3658"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="INFORMATION_PROTECTION_COMPLIANCE";Name="Microsoft 365 E5 Compliance";GUID="184efa21-98c3-4e5d-95ab-d07053a96e67"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="IDENTITY_THREAT_PROTECTION";Name="Microsoft 365 E5 Security";GUID="26124093-3d78-432b-b5dc-48bf992543d5"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="IDENTITY_THREAT_PROTECTION_FOR_EMS_E5";Name="Microsoft 365 E5 Security for EMS E5";GUID="44ac31e7-2999-4304-ad94-c948886741d4"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="M365_F1";Name="Microsoft 365 F1";GUID="44575883-256e-4a79-9da4-ebe9acabe2b2"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="SPE_F1";Name="Microsoft 365 F3";GUID="66b55226-6b4f-492c-910c-a3b7a3c9d993"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="FLOW_FREE";Name="MICROSOFT FLOW FREE";GUID="f30db892-07e9-47e9-837c-80727f46fd3d"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MCOEV";Name="MICROSOFT 365 PHONE SYSTEM";GUID="e43b5b99-8dfb-405f-9987-dc307f34bcbd"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MCOEV_DOD";Name="MICROSOFT 365 PHONE SYSTEM FOR DOD";GUID="d01d9287-694b-44f3-bcc5-ada78c8d953e"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MCOEV_FACULTY";Name="MICROSOFT 365 PHONE SYSTEM FOR FACULTY";GUID="d979703c-028d-4de5-acbf-7955566b69b9"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MCOEV_GOV";Name="MICROSOFT 365 PHONE SYSTEM FOR GCC";GUID="a460366a-ade7-4791-b581-9fbff1bdaa85"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MCOEV_GCCHIGH";Name="MICROSOFT 365 PHONE SYSTEM FOR GCCHIGH";GUID="7035277a-5e49-4abc-a24f-0ec49c501bb5"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MCOEVSMB_1";Name="MICROSOFT 365 PHONE SYSTEM FOR SMALL AND MEDIUM BUSINESS";GUID="aa6791d3-bb09-4bc2-afed-c30c3fe26032"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MCOEV_STUDENT";Name="MICROSOFT 365 PHONE SYSTEM FOR STUDENTS";GUID="1f338bbc-767e-4a1e-a2d4-b73207cc5b93"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MCOEV_TELSTRA";Name="MICROSOFT 365 PHONE SYSTEM FOR TELSTRA";GUID="ffaf2d68-1c95-4eb3-9ddd-59b81fba0f61"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MCOEV_USGOV_DOD";Name="MICROSOFT 365 PHONE SYSTEM_USGOV_DOD";GUID="b0e7de67-e503-4934-b729-53d595ba5cd1"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MCOEV_USGOV_GCCHIGH";Name="MICROSOFT 365 PHONE SYSTEM_USGOV_GCCHIGH";GUID="985fcb26-7b94-475b-b512-89356697be71"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="WIN_DEF_ATP";Name="Microsoft Defender Advanced Threat Protection";GUID="111046dd-295b-4d6d-9724-d52ac90bd1f2"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="CRMPLAN2";Name="MICROSOFT DYNAMICS CRM ONLINE BASIC";GUID="906af65a-2970-46d5-9b58-4e9aa50f0657"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="CRMSTANDARD";Name="MICROSOFT DYNAMICS CRM ONLINE";GUID="d17b27af-3f49-4822-99f9-56a661538792"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="IT_ACADEMY_AD";Name="MS IMAGINE ACADEMY";GUID="ba9a34de-4489-469d-879c-0f0f145321cd"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="TEAMS_FREE";Name="MICROSOFT TEAM (FREE)";GUID="16ddbbfc-09ea-4de2-b1d7-312db6112d70"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="ENTERPRISEPREMIUM_FACULTY";Name="Office 365 A5 for faculty";GUID="a4585165-0533-458a-97e3-c400570268c4"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="ENTERPRISEPREMIUM_STUDENT";Name="Office 365 A5 for students";GUID="ee656612-49fa-43e5-b67e-cb1fdf7699df"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="EQUIVIO_ANALYTICS";Name="Office 365 Advanced Compliance";GUID="1b1b1f7a-8355-43b6-829f-336cfccb744c"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="ATP_ENTERPRISE";Name="Office 365 Advanced Threat Protection (Plan 1)";GUID="4ef96642-f096-40de-a3e9-d83fb2f90211"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="STANDARDPACK";Name="OFFICE 365 E1";GUID="18181a46-0d4e-45cd-891e-60aabd171b4e"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="STANDARDWOFFPACK";Name="OFFICE 365 E2";GUID="6634e0ce-1a9f-428c-a498-f84ec7b8aa2e"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="ENTERPRISEPACK";Name="OFFICE 365 E3";GUID="6fd2c87f-b296-42f0-b197-1e91e994b900"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="DEVELOPERPACK";Name="OFFICE 365 E3 DEVELOPER";GUID="189a915c-fe4f-4ffa-bde4-85b9628d07a0"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="ENTERPRISEPACK_USGOV_DOD";Name="Office 365 E3_USGOV_DOD";GUID="b107e5a3-3e60-4c0d-a184-a7e4395eb44c"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="ENTERPRISEPACK_USGOV_GCCHIGH";Name="Office 365 E3_USGOV_GCCHIGH";GUID="aea38a85-9bd5-4981-aa00-616b411205bf"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="ENTERPRISEWITHSCAL";Name="OFFICE 365 E4";GUID="1392051d-0cb9-4b7a-88d5-621fee5e8711"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="ENTERPRISEPREMIUM";Name="OFFICE 365 E5";GUID="c7df2760-2c81-4ef7-b578-5b5392b571df"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="ENTERPRISEPREMIUM_NOPSTNCONF";Name="OFFICE 365 E5 WITHOUT AUDIO CONFERENCING";GUID="26d45bd9-adf1-46cd-a9e1-51e9a5524128"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="DESKLESSPACK";Name="OFFICE 365 F1";GUID="4b585984-651b-448a-9e53-3b10f069cf7f"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="DESKLESSPACK";Name="OFFICE 365 F3";GUID="4b585984-651b-448a-9e53-3b10f069cf7f"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MIDSIZEPACK";Name="OFFICE 365 MIDSIZE BUSINESS";GUID="04a7fb0d-32e0-4241-b4f5-3f7618cd1162"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="LITEPACK";Name="OFFICE 365 SMALL BUSINESS";GUID="bd09678e-b83c-4d3f-aaba-3dad4abd128b"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="LITEPACK_P2";Name="OFFICE 365 SMALL BUSINESS PREMIUM";GUID="fc14ec4a-4169-49a4-a51e-2c852931814b"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="WACONEDRIVESTANDARD";Name="ONEDRIVE FOR BUSINESS (PLAN 1)";GUID="e6778190-713e-4e4f-9119-8b8238de25df"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="WACONEDRIVEENTERPRISE";Name="ONEDRIVE FOR BUSINESS (PLAN 2)";GUID="ed01faf2-1d88-4947-ae91-45ca18703a96"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="POWERAPPS_PER_USER";Name="POWER APPS PER USER PLAN";GUID="b30411f5-fea1-4a59-9ad9-3db7c7ead579"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="POWER_BI_STANDARD";Name="POWER BI (FREE)";GUID="a403ebcc-fae0-4ca2-8c8c-7a907fd6c235"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="POWER_BI_ADDON";Name="POWER BI FOR OFFICE 365 ADD-ON";GUID="45bc2c81-6072-436a-9b0b-3b12eefbc402"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="POWER_BI_PRO";Name="POWER BI PRO";GUID="f8a1db68-be16-40ed-86d5-cb42ce701560"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="PROJECTCLIENT";Name="PROJECT FOR OFFICE 365";GUID="a10d5e58-74da-4312-95c8-76be4e5b75a0"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="PROJECTESSENTIALS";Name="PROJECT ONLINE ESSENTIALS";GUID="776df282-9fc0-4862-99e2-70e561b9909e"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="PROJECTPREMIUM";Name="PROJECT ONLINE PREMIUM";GUID="09015f9f-377f-4538-bbb5-f75ceb09358a"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="PROJECTONLINE_PLAN_1";Name="PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT";GUID="2db84718-652c-47a7-860c-f10d8abbdae3"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="PROJECTPROFESSIONAL";Name="PROJECT ONLINE PROFESSIONAL";GUID="53818b1b-4a27-454b-8896-0dba576410e6"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="PROJECTONLINE_PLAN_2";Name="PROJECT ONLINE WITH PROJECT FOR OFFICE 365";GUID="f82a60b8-1ee3-4cfb-a4fe-1c6a53c2656c"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="SHAREPOINTSTANDARD";Name="SHAREPOINT ONLINE (PLAN 1)";GUID="1fc08a02-8b3d-43b9-831e-f76859e04e1a"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="SHAREPOINTENTERPRISE";Name="SHAREPOINT ONLINE (PLAN 2)";GUID="a9732ec9-17d9-494c-a51c-d6b45b384dcb"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MCOIMP";Name="SKYPE FOR BUSINESS ONLINE (PLAN 1)";GUID="b8b749f8-a4ef-4887-9539-c95b1eaa5db7"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MCOSTANDARD";Name="SKYPE FOR BUSINESS ONLINE (PLAN 2)";GUID="d42c793f-6c78-4f43-92ca-e8f6a02b035f"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MCOPSTN2";Name="SKYPE FOR BUSINESS PSTN DOMESTIC AND INTERNATIONAL CALLING";GUID="d3b4fe1f-9992-4930-8acb-ca6ec609365e"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MCOPSTN1";Name="SKYPE FOR BUSINESS PSTN DOMESTIC CALLING";GUID="0dab259f-bf13-4952-b7f8-7db8f131b28d"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="MCOPSTN5";Name="SKYPE FOR BUSINESS PSTN DOMESTIC CALLING (120 Minutes)";GUID="54a152dc-90de-4996-93d2-bc47e670fc06"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="VISIOONLINE_PLAN1";Name="VISIO ONLINE PLAN 1";GUID="4b244418-9658-4451-a2b8-b5e2b364e9bd"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="VISIOCLIENT";Name="VISIO Online Plan 2";GUID="c5928f49-12ba-48f7-ada3-0d743a3601d5"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="WIN10_PRO_ENT_SUB";Name="WINDOWS 10 ENTERPRISE E3";GUID="cb10e6cd-9da4-4992-867b-67546b1db821"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="WIN10_VDA_E5";Name="Windows 10 Enterprise E5";GUID="488ba24a-39a9-4473-8ee5-19291e71b002"}
    $listOfAllLicenseSKUNames+=[PSCustomObject]@{SKUId="WINDOWS_STORE";Name="WINDOWS STORE FOR BUSINESS";GUID="6470687e-a428-4b7a-bef2-8a291ad947c9"}
    #endregion SKUNAMES

    $VerbosePreference="Continue"
    write-host "started $(get-date -Format 'HH:mm:ss')"
    $SKUs=get-LicenseSKUs
    $licenseGroups=get-LicenseGroups
    $userCounter=0
    $userLicenses=@()
}

process {
    if($PSCmdlet.ParameterSetName -eq 'UPN') {
        write-verbose "check $UserPrincipalName licenses..." 
        try {
            $msolUser=get-msoluser -UserPrincipalName $UserPrincipalName -ErrorAction Stop
        } catch {
            $_.Exception
            return $null
        }
    }
    if([string]::IsNullOrEmpty( $MSOLUser) ) { 
        Write-Host -ForegroundColor Red "object null."
        return $null
    }

    $userCounter++

    if($msolUser.isLicensed) {
        foreach($license in $msolUser.licenses){
            $exportLicenceObject = [PSCustomObject]@{
                userPrincipalName = $msolUser.userPrincipalName
                AccountSkuId = $license.AccountSkuId
                assignSource = ''
                licenseName  = ''
                licenseGroup = ''
                usageLocation = $MSOLUser.UsageLocation
            }
            foreach($sku in $SKUs) {
                if($license.AccountSkuId -eq $sku.AccountSkuId) {
                    $exportLicenceObject.licenseName = [string]($listOfAllLicenseSKUNames|? SKUId -eq $sku.SKUPartNumber|Select-Object -ExpandProperty name)
                    #if license name is not on the list just put SKUPartNumber
                    if( [string]::IsNullOrEmpty($exportLicenceObject.licenseName) ) {
                        $exportLicenceObject.licenseName=$sku.SKUPartNumber
                    }
                }
            }
            if($license.GroupsAssigningLicense.count -eq 0 -or $license.GroupsAssigningLicense.guid -ieq $msolUser.objectID) {
                $exportLicenceObject.assignSource = 'direct'
            } else {
                $exportLicenceObject.assignSource = 'Group'
                foreach($gbl in $licenseGroups) {
                    if($license.GroupsAssigningLicense.guid -eq $gbl.ObjectId.guid) {
                        $exportLicenceObject.licenseGroup = $gbl.DisplayName
                    }
                }
            }
            if($includeServicePlans) {
                Add-Member -InputObject $exportLicenceObject -NotePropertyName 'servicePlans' -NotePropertyValue ''
                $splans=@()
                $license.ServiceStatus|%{
                    $splans+='['+$_.ServicePlan.ServiceName+':'+$_.ServicePlan.ServiceType+':'+$_.ProvisioningStatus+']'
                }
                $exportLicenceObject.servicePlans=$splans
            }
            $userLicenses+=$exportLicenceObject
            if(-not $exportToCSV) {
                $exportLicenceObject
            }
        }
    } else {
        #no license - return object with some empty values
        $exportLicenceObject=[PSCustomObject]@{
            userPrincipalName = $msolUser.userPrincipalName
            AccountSkuId = ''
            assignSource = ''
            licenseName  = ''
            licenseGroup = ''
            usageLocation = $MSOLUser.UsageLocation
        }
        if($includeServicePlans) {
            Add-Member -InputObject $exportLicenceObject -NotePropertyName 'servicePlans' -NotePropertyValue ''
        }
        $userLicenses+=$exportLicenceObject
        if(-not $exportToCSV) {
            $exportLicenceObject
        }
    }
}

end {
    write-host "processed $userCounter user(s)."
    if($exportToCSV) {
        $userLicenses|
            select-object userPrincipalName,AccountSkuId,assignSource,licenseName,licenseGroup,usageLocation,@{N='servicePlans';E={$_.servicePlans}}|
            export-csv -NoTypeInformation -Path "LicenseReport-$(get-date -Format yyMMddHHmm).csv"
        Write-Host "report exported as .\LicenseReport-$(get-date -Format yyMMddHHmm).csv"
    }
    Write-Host -ForegroundColor Green "ended $(get-date -Format 'HH:mm:ss')"
}
