#Requires -Modules Pipeworks
#Requires -Version 4

function Install-BartenderLabelPrintingWebApp {
    Import-Module BartenderLabelPrinting -Force -PassThru | ConvertTo-ModuleService -Force
}

$LabelTypes = [pscustomobject][ordered]@{
    Name = "Tumbler"
    LabelTemplatePath = $env:TumblerLabelTemplatePath
},
[pscustomobject][ordered]@{
    Name="Accessory"
    LabelTemplatePath = $env:AccessoryLabelTemplatePath
}

function Invoke-BartenderLabelPrint {
    param(
        $OrderNumber,
        [ValidateSet("Tumbler","Accessory")]
        $LabelType,

        [ValidateScript({$_ -in  $(Get-ZebraPrinters)})]
        [parameter(Mandatory)]
        $Printer = $(Get-BartenderLabelPrintingDefaultPrinter)
    )

    $LabelTypes
}

function Get-ZebraPrinter {
    Get-Printer -ComputerName disney | 
    where drivername -match ZDesigner 
} 

function Set-DefaultTumblerUPCPrinter {
    param(
        [ValidateScript({
            
        })]
        $PrinterName
    )
}

function Set-BartenderLabelPrintingDefaultPrinter {
    param(
        [ValidateScript({
            $_ -in  $(
                Get-ZebraPrinter |
                select -ExpandProperty name
            )
        })]
        $Printer
    )
    [System.Web.HttpCookie]$MyCookie = [System.Web.HttpCookie] "PrinterName" 
    $MyCookie.Value = "Turtle"    
    $Response.Cookies.Add($MyCookie)

    #Store this printer choice in a cookie
}

function Get-BartenderLabelPrintingDefaultPrinter {
    #Read DefaultPrinterFromCookie
    $request.cookies["PrinterName"].Value
}

function Get-MESCustomerNumberFromOrderNumber {
    param(
        $OrderNumber
    )
    
}

function Get-ThermalPrinter {
    param(
        [validateset("ShippingGS1","ShippingShipLabel","TumblerUPC","AccessoryUPC")]
        [ValidateScript({$_ -in $ThermalPrinterTypes.Name})]
        $ThermalPrinterTypeName
    )
    $ThermalPrinterType = $ThermalPrinterTypes | where name -eq $ThermalPrinterTypeName

    Get-ZebraPrinter |
    Add-PrinterMetadataMember -PassThrough |
    where LabelWidth -EQ $($ThermalPrinterType.LabelWidth) |
    where LabelLength -EQ $($ThermalPrinterType.LabelLength) |
    where DPI -GE $($ThermalPrinterType.MinimumDPI) |
    where MediaType -EQ $($ThermalPrinterType.MediaType)
}

$ThermalPrinterTypes = [pscustomobject][ordered]@{
    Name = "ShippingGS1"
    LabelWidth = 4
    LabelLength = 6
    MinimumDPI = 203
    MediaType = "Direct-Thermal"
},
[pscustomobject][ordered]@{
    Name = "ShippingShipLabel"
    LabelWidth = 4
    LabelLength = 8
    MinimumDPI = 203
    MediaType = "Direct-Thermal"
},
[pscustomobject][ordered]@{
    Name = "TumblerUPC"
    LabelWidth = 1.25
    LabelLength = 1.5
    MinimumDPI = 300
    MediaType = "Thermal-Transfer"
},
[pscustomobject][ordered]@{
    Name = "AccessoryUPC"
    LabelWidth = 2.25
    LabelLength = 2
    MinimumDPI = 300
    MediaType = "Thermal-Transfer"
}


function Add-PrinterMetadataMember {
    param(
        [Switch]$PassThrough,
        [Parameter(ValueFromPipeline, Mandatory)]$Printer
    )
    process {
        $PrinterMetadata = try {$Printer.comment | convertfrom-json} catch {}
        foreach ($Property in $PrinterMetadata.psobject.Properties) {
            $Printer | Add-Member -MemberType NoteProperty -Name $($Property.Name) -Value $($Property.Value)
        }
        if($PassThrough) { $Printer }
    }
}

function Invoke-TervisBartenderCommanderOrderNumber {
    param(
        $PathToLabelFile,
        $Printer,
        $LabelParameters
    )

    $CommanderRequestXML = @"
<?xml version="1.0" encoding="utf-8"?><XMLScript Version="2.0">
<Command>
<Print ReturnPrintData="true" ReturnLabelData="true" ReturnChecksum="false">
<Format>$PathToLabelFile</Format>
<PrintSetup>
<Printer>$Printer</Printer>
<Performance><AllowFormatCaching>true</AllowFormatCaching><AllowGraphicsCaching>true</AllowGraphicsCaching><AllowSerialization>true</AllowSerialization><AllowStaticGraphics>true</AllowStaticGraphics><AllowStaticObjects>true</AllowStaticObjects><AllowVariableDataOptimization>false</AllowVariableDataOptimization><WarnWhenUsingTrueTypeFonts>true</WarnWhenUsingTrueTypeFonts></Performance>
</PrintSetup>
<QueryPrompt Name="WorkOrderId"><Value>1042-9080007-1</Value></QueryPrompt><QueryPrompt Name="ContainerCode"><Value>006087120</Value></QueryPrompt>
</Print>
</Command>
</XMLScript>
"@  

}