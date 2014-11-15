<?php
$arrResponse = array(); 
$arrResponse['Test'] = array(); 
$arrResponse['Test'][] = array('type'=>'string','value'=>'abc');
 
$strWorksheetName = "Sheet1"; // Excel worksheet name

$strFileName =  $strWorksheetName."_".date("MdY_hi").".xml"; // download attachment file name
$intDateTypeCell = false;  // flag define that there is a date type node value
$intDateTimeTypeCell = false;  // flag define that there is a date Time type node value
$strRow = ''; // content rows data collector
$strRowHeader = ""; // content header data collector
if(!empty($arrResponse)){        
    $strRowHeader .= "\t<ss:Row>\n"; // open header row
    $arrRowsData = array();
    foreach($arrResponse as $strHeader=>$arrRows){
        if(empty($arrRows) || !is_array($arrRows) ){
            // if this column do not contain needed Excel XML info to be generated
             continue;
        }
        // collect the headers  
        $strHeader = preg_replace("([A-Z])", " $0", $strHeader); // add space before capital letters
        $strRowHeader .= "\t\t<ss:Cell  ss:StyleID=\"s11\"><Data ss:Type=\"String\">{$strHeader}</Data></ss:Cell>\n"; 
        // loop over content    
        foreach($arrRows as $key=>$arrRow){
            $arrRowsData[$key][$strHeader] = $arrRow;             
         }  
    }
    $strRowHeader .= "\t</ss:Row>\n"; // close header row    
    // loop over content rows
    foreach($arrRowsData as $arrRowData){
          $strRow .= "\t<ss:Row>\n";  // open content row
            // loop over row content
            foreach($arrRowData as $arrCell){
                // column value type related setting
                switch($arrCell['type'] ){
                    case "Date":
                        $strTypeStyleID = "ss:StyleID=\"s21\"";
                        $intDateTypeCell = true; // flag that a date type node exist 
                    break;
                    case "DateTime":
                        $strTypeStyleID = "ss:StyleID=\"s22\"";
                        $intDateTimeTypeCell = true; // flag that a date type node exist 
                        $arrCell['value'] = str_replace(' ','T',$arrCell['value']); // force T as separator between date and time
                    break;
                    default:
                        $strTypeStyleID = "";
                    break;
                }
      
                // add content to a cell   
                $strRow .= "\t\t<ss:Cell {$strTypeStyleID}>";   // open content cell
                if(strlen($arrCell['value'])>0){  // if there is any value then add the data portion
                   $strRow .= "<Data ss:Type=\"{$arrCell['type']}\">{$arrCell['value']}</Data>"; // add value to cell
                }
                $strRow .= "</ss:Cell>\n";   // close content cell
            }
          $strRow .= "\t</ss:Row>\n"; // close content row
    }         
    // buld the Excel XML stracture
    $strXML = "<?xml version=\"1.0\"?>\n";
    $strXML .= "<?mso-application progid=\"Excel.Sheet\"?>\n";
    $strXML .= "<Workbook\n";
    $strXML .= "\txmlns:x=\"urn:schemas-microsoft-com:office:excel\"";
    $strXML .= "\txmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"";
    $strXML .= "\txmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">\n";
    $strXML .= "<Styles>\n";
    $strXML .= "\t<Style ss:ID=\"Default\" ss:Name=\"Normal\">\n";
    $strXML .= "\t<Alignment ss:Vertical=\"Bottom\"/>\n";
    $strXML .= "\t<Borders/>\n";
    $strXML .= "\t<ss:Font ss:Size=\"12\" ss:FontName=\"Tahoma\" ss:Color=\"#5F772E\"/>\n"; // define page font
    $strXML .= "\t<Interior/>\n";
    $strXML .= "\t<NumberFormat/>\n";
    $strXML .= "\t<Protection/>\n";
    $strXML .= "</Style>\n";
    // header style
    $strXML .= "<Style ss:ID=\"s11\">\n";
    $strXML .= "\t<Font x:Family=\"Swiss\" ss:Color=\"#5F772E\" ss:Bold=\"1\"/>\n";
    $strXML .= "</Style>\n";  
    if($intDateTimeTypeCell){
        // if there is a date type node value
        // define the date format
        $strXML .= "<Style ss:ID=\"s22\">\n";
        $strXML .= "\t<NumberFormat ss:Format=\"yyyy\-mm\-dd\ hh:mm:ss\"/>\n";
        $strXML .= "</Style>\n";   
    } 
    if($intDateTypeCell){
        // if there is a date type node value
        // define the date format
        $strXML .= "<Style ss:ID=\"s21\">\n";
        $strXML .= "\t<NumberFormat ss:Format=\"yyyy\-mm\-dd\"/>\n";
        $strXML .= "</Style>\n";   
    }  
    $strXML .= "</Styles>\n";    
    $strXML .= "\t<Worksheet ss:Name=\"{$strWorksheetName}\">\n";
    $strXML .= "<ss:Table>\n";
    // add the header row
    $strXML .= $strRowHeader;
    // add the content data rows
    $strXML .= $strRow;
    $strXML .= "</ss:Table>\n";
    $strXML .= "</Worksheet>\n";
    $strXML .= "</Workbook>\n";  
    // generate browser download headers
    header('Content-Encoding: UTF-8');
    header("Pragma: public");
    header("Expires: 0");
   // header("Content-type: application/octet-stream");
   // set default open method to Excel
    header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); 
    // set the downloadeble file new name  
    header("Content-Disposition: attachment; filename=\"{$strFileName}\"");  
    echo $strXML;
    exit();            
}

   
?>
