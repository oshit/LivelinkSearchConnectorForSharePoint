# Livelink OpenSearch Connector OSDX File Generator
#
# Version 1.0
#
# Creates a new file descriptor for the Livelink OpenSearch Connector.
# The created OSDX file will be bound to a specific ASP.NET web site
# (running the connector) and to a specific Livelink server.
#
# OSDX files can be imported by OpenSearch clients which can execute
# search queries against servers compliant with the OpenSearch 1.1
# specification. Windows Explorer is an OpenSearch client, for example.
#
# This script expects an ASP.NET web site deployed by Livelink OpenSearch
# Connector for SharePoint 2013 Preview or other compatible one.
#
# Usage example:
#   New-LivelinkSearchConnectorOSDX.ps1 c:\localhost.osdx
#     http://localhost/_layouts/LivelinkOpenSearch
#     http://localhost/livelink/llisapi.dll -TargetAppID livelink
#
# Copyright © 2012 Ferdinand Prantl <prantlf@gmail.com>
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

[CmdletBinding(SupportsShouldProcess = $true)]

Param(
    [Parameter(
     HelpMessage = 'Path to the OSDX file to be created, including the file name. Example: c:\localhost.osdx')]
    [Alias('o')]
    [string] $OutputFile,

    [Parameter(Mandatory = $true,
     HelpMessage = 'URL of the Livelink OpenSearch Connector web site. Example: http://localhost/_layouts/LivelinkOpenSearch')]
    [ValidatePattern('https?://.+')]
    [string] $ConnectorURL,

    [Parameter(Mandatory = $true,
     HelpMessage = 'URL of the Livelink CGI exectable. Example: http://localhost/livelink/llisapi.dll')]
    [ValidatePattern('https?://.+')]
    [string] $LivelinkURL,

    [Parameter(Mandatory = $true, ParameterSetName = "LivelinkWithoutSSO",
     HelpMessage = 'Identifier of a target application in the SharePoint Secure Store Service which contains credentials of a Livelink system administrator. Example: livelink')]
    [string] $TargetAppID,

    [Parameter(Mandatory = $true, ParameterSetName = "LivelinkWithoutSSO",
     HelpMessage = 'Computes Livelink user login name from the login name of current (Windows or ASP.NET) user. Supported placeholders are login, user and domain. Supported modifiers are :lc and :uc. Example: {user:lc}')]
    [string] $LoginPattern,

    [Parameter(Mandatory = $true, ParameterSetName = "LivelinkWithSSO",
     HelpMessage = 'Enables SSO of the current user for the Livelink search URL requests.')]
    [switch] $UseSSO,

    [Parameter(
     HelpMessage = 'URL query parameters to append to the Livelink search URL. Default: lookfor1=allwords&fullTextMode=allwords&hhterms=true')]
    [switch] $ExtraParams,

    [Parameter(HelpMessage = 'Short name of the search source. Default: Search Enterprise at <hostname>')]
    [string] $ShortName,

    [Parameter(HelpMessage = 'Long name of the search source. Default: Search Livelink Enterprise Workspace at <hostname>')]
    [string] $LongName,

    [Parameter(HelpMessage = 'Internal identifier of the search source. Default: search_<hostname>')]
    [string] $InternalName,

    [Parameter(HelpMessage = 'Description of the search source. Default: Searches content in the Enterprise Workspace of the Livelink server at <hostname>.')]
    [string] $Description,

    [Parameter(HelpMessage = 'Accepts HTTPS responses from the Livelink server with invalid SSL certificate.')]
    [switch] $IgnoreSSLWarnings,

    [Parameter(HelpMessage = 'Suppresses printing the version and license information on the console when the script executions starts.')]
    [Alias('q')]
    [switch] $NoLogo
)

if (!$NoLogo) {
    Write-Host "
Livelink OpenSearch Connector OSDX File Generator 1.0

This program comes with ABSOLUTELY NO WARRANTY; for details type
'Get-Help  New-LivelinkSearchConnectorOSDX'.
This is free software, and you are welcome to redistribute it
under certain conditions; see the note above for details.
"
}

# Enable calling System.Web.HttpUtility.UrlEncode below.
$systemWeb = [Reflection.Assembly]::LoadWithPartialName('System.Web')

if ($LivelinkURL -match 'https?://([^/:]+)') {
    $hostName = $Matches[1]
}
if ($LivelinkURL -match 'https?://[^/]+') {
    $urlBase = $Matches[0]
}

$ConnectorURL = $ConnectorURL.TrimEnd('/')

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
if (!$ExtraParams -or $ExtraParams -eq '') {
    $params = 'lookfor1=allwords&fullTextMode=allwords&hhterms=true'
    Write-Verbose "ExtraParams: $ExtraParams"
	$params = [Web.HttpUtility]::UrlEncode($params)
    $params += '&'
} else {
    $params = ''
}

if ($UseSSO) {
    $authentication = "useSSO=true"
} else {
    $authentication = "targetAppID=$TargetAppID&loginPattern=" +
        [Web.HttpUtility]::UrlEncode($LoginPattern)
}
if ($IgnoreSSLWarnings) {
    $certification = 'ignoreSSLWarnings=true&'
} else {
    $certification = ''
}
$urlTemplate = "$ConnectorURL/ExecuteQuery.aspx?query={searchTerms}&" +
    "livelinkUrl=$([Web.HttpUtility]::UrlEncode($LivelinkURL))&$authentication&" +
    "count={count}&startIndex={startIndex}&extraParams=$params&" +
    "maxSummaryLength=185&$($certification)inputEncoding={inputEncoding}&" +
    "outputEncoding={outputEncoding}&language={language}"
Write-Verbose "SearchURLTemplate: $urlTemplate"
$urlTemplate = $urlTemplate -replace '&', '&amp;'

$content = @"
<?xml version="1.0" encoding="UTF-8"?>
<OpenSearchDescription xmlns="http://a9.com/-/spec/opensearch/1.1/">
  <ShortName>$ShortName</ShortName>
  <LongName>$LongName</LongName>
  <InternalName xmlns="http://schemas.microsoft.com/Search/2007/location">$InternalName</InternalName>
  <Description>$Description</Description>
  <Image height="32" width="32" type="image/png">$urlBase/img/style/images/app_content_server32_b8.png</Image>
  <Url type="application/rss+xml" template="$urlTemplate"/>
  <Url type="text/html" template="$urlTemplate&amp;format=html"/>
  <Developer>Ferdinand Prantl</Developer>
  <Contact>prantlf@gmail.com</Contact>
  <Attribution>Copyright (c) 2012 Ferdinand Prantl, All Rights Reserved</Attribution>
  <SyndicationRight>Open</SyndicationRight>
  <InputEncoding>UTF-8</InputEncoding>
  <OutputEncoding>UTF-8</OutputEncoding>
  
  <ms-ose:ResultsProcessing format="application/rss+xml" xmlns:ms-ose="http://schemas.microsoft.com/opensearchext/2009/">
    <ms-ose:PropertyDefaultValues>
      <ms-ose:Property schema="http://schemas.microsoft.com/windows/2008/propertynamespace" name="System.PropList.ContentViewModeForSearch">prop:~System.ItemNameDisplay;System.LayoutPattern.PlaceHolder;~System.ItemPathDisplay;~System.Search.AutoSummary;System.LayoutPattern.PlaceHolder;System.LayoutPattern.PlaceHolder;System.LayoutPattern.PlaceHolder</ms-ose:Property>
    </ms-ose:PropertyDefaultValues>
  </ms-ose:ResultsProcessing>
</OpenSearchDescription>
"@

if ($OutputFile) {
    Write-Verbose "Writing the OpenSearch descriptor..."
    $outFile = Out-File -FilePath $OutputFile -InputObject $content
    Write-Verbose "Done."
} else {
    Write-Output $content
}
