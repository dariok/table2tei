xquery version "3.0";

let $filter := function ($path as xs:string, $data-type as xs:string, $param as item()*) as xs:boolean {
    switch ($path)
        case "xl/worksheets/sheet1.xml"
        case "xl/sharedStrings.xml"
        case "content.xml"
            return true()
        default
            return false()
}

let $process := function ($path as xs:string, $data-type as xs:string, $data as item()?, $param as item()*) {
  <pkg:part pkg:name="/{$path}" xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
    <pkg:xmlData>{
      $data
    }</pkg:xmlData>
  </pkg:part>
}

let $title := if (ends-with(request:get-uploaded-file-name('file'), '.ods'))
	then substring-before(request:get-uploaded-file-name('file'), '.ods')
	else substring-before(request:get-uploaded-file-name('file'), '.xslx')

let $origFileData := string(request:get-uploaded-file-data('file'))
let $incoming := 
  <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">{
    compression:unzip($origFileData, $filter, (), $process, ())
  }</pkg:package>

let $type := if (local-name($incoming/*[1]) = 'document-content')
	then 'OO'
	else 'MS'
let $xslt := if ($type = 'OO')
	then doc('xslt/oo2tei.xsl')
	else doc('xslt/ms2tei.xsl')
let $add := request:get-parameter('xslt', 'none')
let $post := request:get-parameter('firstheading', false())

let $params :=
	<parameters>
		<param name="heading" value="true" />
		<param name="title" value="{$title}" />
		<param name="filename" value="{request:get-uploaded-file-name('file')}" />
	</parameters>

let $firstPass :=
	if ($post)
		then transform:transform($incoming, $xslt, $params)
		else transform:transform($incoming, $xslt, ())

let $result := if ($add != 'none')
	then transform:transform($firstPass, doc($add), ())
	else $firstPass

let $filename := $title || '.xml'
let $header := response:set-header("Content-Disposition", concat("attachment; filename=", $filename))
return $result
