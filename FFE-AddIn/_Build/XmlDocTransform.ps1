param (
    [string] $XmlFile,
    [string] $XdtFile,
    [string] $XmlFileNew
);

function XmlDocTransform($xml, $xdt, $xmlNew)
{
    if (!$xml -or !(Test-Path -path $xml -PathType Leaf)) {
        throw "XML file not found. $xml";
    }
    if (!$xdt -or !(Test-Path -path $xdt -PathType Leaf)) {
        throw "XDT file not found. $xdt";
    }

    $scriptPath = (Get-Variable MyInvocation -Scope 1).Value.InvocationName | split-path -parent
    Add-Type -LiteralPath "$scriptPath\Microsoft.Web.XmlTransform.dll"

    $xmlDoc = New-Object Microsoft.Web.XmlTransform.XmlTransformableDocument;
    $xmlDoc.PreserveWhitespace = $true
    $xmlDoc.Load($xml);

    $transf = New-Object Microsoft.Web.XmlTransform.XmlTransformation($xdt);
    if ($transf.Apply($xmlDoc) -eq $false)
    {
        throw "Transformation failed."
    }
    $xmlDoc.Save($xmlNew);
}

XmlDocTransform -xml $XmlFile -xdt $XdtFile -xmlNew $XmlFileNew;