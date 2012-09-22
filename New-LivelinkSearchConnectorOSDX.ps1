# Livelink OpenSearch Connector OSDX File Generator
#
# Version 1.0
#
# Creates an OpenSearch 1.1 file descriptor (OSDX) for the Livelink
# OpenSearch Connector for SharePoint 2013 Preview.
#
# See the description in the Get-Help supporting comment below.
#
# Copyright © 2012 Ferdinand Prantl <prantlf@gmail.com>
# All rights reserved.
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

<#
.SYNOPSIS
    Creates an OpenSearch 1.1 file descriptor (OSDX) for the Livelink
    OpenSearch Connector for SharePoint 2013 Preview.
.DESCRIPTION
    Creates a new file descriptor for the Livelink OpenSearch Connector.
    The created OSDX file will be bound to a specific ASP.NET web site
    (running the connector) and to a specific Livelink server. (This
    script expects an ASP.NET web site deployed by Livelink OpenSearch
    Connector for SharePoint 2013 Preview or other compatible one.)
    Default parameters are used if no or empty value is provided. If you
    want to pass an empty string as a value, use whitespace. (Parameter
    values are trimmed.)

    OSDX files can be imported by OpenSearch clients which can execute
    search queries against servers compliant with the OpenSearch 1.1
    specification. Windows Explorer is an OpenSearch client, for example.
.PARAMETER OutFile
    Path to the OSDX file to be created, including the file name. If
    omitted the content will be written to the pipe output.
    Example: c:\localhost.osdx.
.PARAMETER ConnectorURL
    URL of the Livelink OpenSearch Connector web site; without the ASPX
    page. Example: http://localhost/_layouts/LivelinkOpenSearch.
.PARAMETER LivelinkURL
    URL of the Livelink CGI executable which is used to open the Livelink
    web UI. Example: http://localhost/livelink/llisapi.dll.
.PARAMETER TargetAppID
    Identifier of a target application in the SharePoint Secure Store
    which contains credentials of a Livelink system administrator.
    This parameter requires providing LoginPattern and omitting UseSSO.
    Example: livelink.
.PARAMETER LoginPattern
    Template of the Livelink user login name. The following placeholders
    are resolved using the login name of current (Windows or ASP.NET)
    user: login, user and domain. Supported modifiers are :lc and :uc.
    This parameter requires providing TargetAppID and omitting UseSSO.
    Example: {user:lc}.
.PARAMETER UseSSO
    Enables SSO of the current user for the Livelink search URL requests.
    This parameter cannot be used together with TargetAppID and
    LoginPattern; they are for connections without SSO.
.PARAMETER ExtraParams
    Additional URL query parameters to append to the Livelink search URL.
    Default: lookfor1=allwords&fullTextMode=allwords&hhterms=true.
.PARAMETER ShortName
    Short name of the search source.
    Default: Search Enterprise at <hostname>.
.PARAMETER LongName
    Long name of the search source.
    Default: Search Livelink Enterprise Workspace at <hostname>.
.PARAMETER InternalName
    Internal identifier of the search source. Default: search_<hostname>.
.PARAMETER Description
    Description of the search source. Default: Searches content in the
    Enterprise Workspace of the Livelink server at <hostname>.
.PARAMETER IgnoreSSLWarnings
    Accepts HTTPS responses from a Livelink server with invalid SSL
    certificate.
.PARAMETER MaxSummaryLength
    Limits the maximum length of the textual summary that is displayed
    below a search hit to give a hint what is the document about. If less
    than zero the text is not limited. Default: 185.
.PARAMETER ReportErrorAsHit
    Errors occurring during the search are reported by HTTP error code 500
    by default. This parameter makes the error message be returned as
    single search hit for OpenSearch clients that do not show HTTP errors
    to the user.
.PARAMETER NoLogo
    Suppresses printing the version and license information on the console
    when the script executions starts.
.PARAMETER WhatIf
    Prints an information that the OSDX descriptor would be written to a file
    instead of actually creating the file if the option OutFile is used.
.PARAMETER Confirm
    Prompts for a confirmation that the OSDX descriptor would be written
    to a file before actually creating the file if the option OutFile is used.
.INPUTS
    The input is provided by script parameters.
.OUTPUTS
    The output consists of the OSDX file descriptor. Its content can be
    either written to a file or printed on the console.
.EXAMPLE
    New-LivelinkSearchConnectorOSDX.ps1 -OutputFile c:\localhost.osdx -ConnectorURL http://localhost/_layouts/LivelinkOpenSearch -LivelinkURL http://localhost/livelink/llisapi.dll -TargetAppID livelink -LoginPattern {user:lc}

    Writes an OSDX file at c:\localhost.osdx executing queries against
    http://localhost/_layouts/LivelinkOpenSearch/ExecuteQuery.aspx which
    will connect to Livelink at http://localhost/livelink/llisapi.dll.
    The search will be authorized by Livelink system administrator
    credentials stored in the SharePoint Secure Store and impersonated
    for the current user by using his user name in lower-case.
.EXAMPLE
    New-LivelinkSearchConnectorOSDX.ps1 -UseSSO -ConnectorURL http://localhost/_layouts/LivelinkOpenSearch -LivelinkURL http://localhost/livelink/llisapi.dll

    Prints (on the console ) the OSDX file content to execute queries
    against http://localhost/_layouts/LivelinkOpenSearch/ExecuteQuery.aspx
    which will connect to Livelink at http://localhost/livelink/llisapi.dll.
    Authenticatio nof the current user with the search service will
    depend on SSO configured between the connector and Livelink servers.
.NOTES
    Version:   1.0
    Date:      September 8, 2012
    Author:    Ferdinand Prantl <prantlf@gmail.com>
    Copyright: © 2012 Ferdinand Prantl, All rights reserved.
    License:   GPL
.LINK
    http://prantlf.blogspot.com
    http://github.com/prantlf/LivelinkSearchConnectorForSharePoint.git
#>

# Support Verbose, WhatIf and Confirm switches.
[CmdletBinding(SupportsShouldProcess = $true)]

Param(
    [Parameter(
     HelpMessage = 'Path to the OSDX file to be created, including the file name. If omitted the content will be written to the pipe output. Example: c:\localhost.osdx')]
    [Alias('o')]
    [string] $OutputFile,

    [Parameter(Mandatory = $true,
     HelpMessage = 'URL of the Livelink OpenSearch Connector web site; without the ASPX page. Example: http://localhost/_layouts/LivelinkOpenSearch')]
    [ValidatePattern('https?://.+')]
    [string] $ConnectorURL,

    [Parameter(Mandatory = $true,
     HelpMessage = 'URL of the Livelink CGI executable which is used to open the Livelink web UI. Example: http://localhost/livelink/llisapi.dll')]
    [ValidatePattern('https?://.+')]
    [string] $LivelinkURL,

    [Parameter(Mandatory = $true, ParameterSetName = "LivelinkWithoutSSO",
     HelpMessage = 'Identifier of a target application in the SharePoint Secure Store which contains credentials of a Livelink system administrator. Example: livelink')]
    [string] $TargetAppID,

    [Parameter(Mandatory = $true, ParameterSetName = "LivelinkWithoutSSO",
     HelpMessage = 'Template of the Livelink user login name. The following placeholders are resolved using the login name of current (Windows or ASP.NET) user: login, user and domain. Supported modifiers are :lc and :uc. Example: {user:lc}')]
    [string] $LoginPattern,

    [Parameter(Mandatory = $true, ParameterSetName = "LivelinkWithSSO",
     HelpMessage = 'Enables SSO of the current user for the Livelink search URL requests.')]
    [switch] $UseSSO,

    [Parameter(
     HelpMessage = 'Additional URL query parameters to append to the Livelink search URL. Default: lookfor1=allwords&fullTextMode=allwords&hhterms=true')]
    [string] $ExtraParams,

    [Parameter(HelpMessage = 'Short name of the search source. Default: Search Enterprise at <hostname>')]
    [string] $ShortName,

    [Parameter(HelpMessage = 'Long name of the search source. Default: Search Livelink Enterprise Workspace at <hostname>')]
    [string] $LongName,

    [Parameter(HelpMessage = 'Internal identifier of the search source. Default: search_<hostname>')]
    [string] $InternalName,

    [Parameter(HelpMessage = 'Description of the search source. Default: Searches content in the Enterprise Workspace of the Livelink server at <hostname>.')]
    [string] $Description,

    [Parameter(HelpMessage = 'Accepts HTTPS responses from a Livelink server with invalid SSL certificate.')]
    [switch] $IgnoreSSLWarnings,

    [Parameter(HelpMessage = 'Limits the maximum length of the textual summary that is displayed below a search hit to give a hint what is the document about. If less than zero the text is not limited. Default: 185')]
    [int] $MaxSummaryLength,

    [Parameter(HelpMessage = 'Makes the error message be returned as single search hit instead of returnung HTTP error code 500.')]
    [switch] $ReportErrorAsHit,

    [Parameter(HelpMessage = 'Suppresses printing the version and license information on the console when the script executions starts.')]
    [Alias('q')]
    [switch] $NoLogo
)

if (!$NoLogo) {
    Write-Host "
Livelink OpenSearch Connector OSDX File Generator 1.0

This program comes with ABSOLUTELY NO WARRANTY; for details type
'Get-Help New-LivelinkSearchConnectorOSDX'.
This is free software, and you are welcome to redistribute it
under certain conditions; see the note above for details.
"
}

# Enable calling UrlEncode and HtmlEncode from System.Web.HttpUtility below.
$systemWeb = [Reflection.Assembly]::LoadWithPartialName('System.Web')

# Valid URL has been checked by the parameter validator above.
[void] ($LivelinkURL -match 'https?://([^/:]+)')
$hostName = $Matches[1]
[void] ($LivelinkURL -match 'https?://[^/]+');
$urlBase = $Matches[0]

# Setting default parameter values if no values were passed to the script.
if (!$ShortName) {
    $ShortName = "Enterprise at $hostName"
    Write-Verbose "ShotName: $ShortName"
}
if (!$LongName) {
    $LongName = "Search Livelink Enterprise Workspace at $hostName"
    Write-Verbose "LongName: $LongName"
}
if (!$InternalName) {
    $InternalName = "search_$hostName"
    Write-Verbose "InternalName: $InternalName"
}
if (!$Description) {
    $Description = "Searches content in the Enterprise Workspace of the Livelink server at $hostName."
    Write-Verbose "Description: $Description"
}
if (!$ExtraParams) {
    $ExtraParams = 'lookfor1=allwords&fullTextMode=allwords&hhterms=true'
    Write-Verbose "ExtraParams: $ExtraParams"
}
if (!$MaxSummaryLength) {
    $MaxSummaryLength = 185
    Write-Verbose "MaxSummaryLength: $MaxSummaryLength"
}

# Normalize parameter values to fix possible user input errors.
$ConnectorURL = $ConnectorURL.Trim().TrimEnd('/')
$LivelinkURL = $LivelinkURL.Trim()
$ShortName = $ShortName.Trim()
$LongName = $LongName.Trim()
$ShortName = $ShortName.Trim()
$InternalName = $InternalName.Trim()
$Description = $Description.Trim()
$ExtraParams = $ExtraParams.Trim().Trim('&')

# Prepare the search template URL.
if ($UseSSO) {
    $authentication = "useSSO=true"
} else {
    $authentication = "targetAppID=" + [Web.HttpUtility]::UrlEncode($TargetAppID) +
        "&loginPattern=" + [Web.HttpUtility]::UrlEncode($LoginPattern)
}
if ($IgnoreSSLWarnings) {
    $certification = 'ignoreSSLWarnings=true&'
} else {
    $certification = ''
}
if ($MaxSummaryLength -gt 0) {
    $limit = "maxSummaryLength=$MaxSummaryLength&"
} else {
    $limit = ''
}
if ($ReportErrorAsHit) {
    $error = "reportErrorAsHit=true&"
} else {
    $error = ''
}
$urlTemplate = "$ConnectorURL/ExecuteQuery.aspx?query={searchTerms}&" +
    "livelinkUrl=$([Web.HttpUtility]::UrlEncode($LivelinkURL))&" +
    "$authentication&count={count}&startIndex={startIndex}&extraParams=" +
    "$([Web.HttpUtility]::UrlEncode($ExtraParams))&$limit$error" +
    "$($certification)inputEncoding={inputEncoding}&" +
    "outputEncoding={outputEncoding}&language={language}"
Write-Verbose "SearchURLTemplate: $urlTemplate"
$urlOSDX = "$ConnectorURL/GetOSDX.aspx?" +
    "livelinkUrl=$([Web.HttpUtility]::UrlEncode($LivelinkURL))&" +
    "$authentication&$limit$error$($certification)extraParams=" +
    "$([Web.HttpUtility]::UrlEncode($ExtraParams))";
Write-Verbose "OSDXURL: $urlOSDX"

# Generate the OSDX file content.
$content = @"
<OpenSearchDescription xmlns="http://a9.com/-/spec/opensearch/1.1/">
  <ShortName>$([Web.HttpUtility]::HtmlEncode($ShortName))</ShortName>
  <LongName>$([Web.HttpUtility]::HtmlEncode($LongName))</LongName>
  <InternalName xmlns="http://schemas.microsoft.com/Search/2007/location">$([Web.HttpUtility]::HtmlEncode($InternalName))</InternalName>
  <Description>$([Web.HttpUtility]::HtmlEncode($Description))</Description>
  <Image height="32" width="32" type="image/png">$([Web.HttpUtility]::HtmlEncode($urlBase))/img/style/images/app_content_server32_b8.png</Image>
  <Url type="application/rss+xml" rel="results" template="$([Web.HttpUtility]::HtmlEncode($urlTemplate))"/>
  <Url type="text/html" template="$([Web.HttpUtility]::HtmlEncode($urlTemplate))&amp;format=html"/>
  <Url type="application/opensearchdescription+xml" rel="self" template="$([Web.HttpUtility]::HtmlEncode($urlOSDX))"/>
  <Tags>Livelink OpenSearch</Tags>
  <Query role="example" searchTerms="livelink" />
  <AdultContent>false</AdultContent>
  <Developer>Ferdinand Prantl</Developer>
  <Contact>prantlf@gmail.com</Contact>
  <Attribution>Copyright (c) 2012 Ferdinand Prantl, All rights reserved.</Attribution>
  <SyndicationRight>Open</SyndicationRight>
  <InputEncoding>UTF-8</InputEncoding>
  <OutputEncoding>UTF-8</OutputEncoding>
  <Language>*</Language>
  
  <ms-ose:ResultsProcessing format="application/rss+xml" xmlns:ms-ose="http://schemas.microsoft.com/opensearchext/2009/">
    <ms-ose:PropertyDefaultValues>
      <ms-ose:Property schema="http://schemas.microsoft.com/windows/2008/propertynamespace" name="System.PropList.ContentViewModeForSearch">prop:~System.ItemNameDisplay;System.LayoutPattern.PlaceHolder;~System.ItemPathDisplay;~System.Search.AutoSummary;System.LayoutPattern.PlaceHolder;System.LayoutPattern.PlaceHolder;System.LayoutPattern.PlaceHolder</ms-ose:Property>
    </ms-ose:PropertyDefaultValues>
  </ms-ose:ResultsProcessing>
</OpenSearchDescription>
"@

# Write the OSDX file or prints its content on the console.
if ($OutputFile) {
    Write-Verbose "Writing the OpenSearch descriptor..."
    $outFile = Out-File -FilePath $OutputFile -InputObject $content
    Write-Verbose "Done."
} else {
    Write-Output $content
}
