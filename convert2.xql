xquery version "3.0";

declare namespace tei	= "http://www.tei-c.org/ns/1.0";

(: OpenOffice :)
declare namespace office="urn:oasis:names:tc:opendocument:xmlns:office:1.0";

(: MS :)
declare namespace excel="http://schemas.openxmlformats.org/spreadsheetml/2006/main";

let $filter := function ($path as xs:string, $data-type as xs:string, $param as item()*) as xs:boolean {
    switch ($path)
        case "xl/worksheets/sheet1.xml"
        case "xl/sharedStrings.xml"
        case "content.xml"
            return true()
        default
            return false()
}

let $entry-data := function ($path as xs:string, $data-type as xs:string, $data as item()?, $param as item()*) {
  $data
}

			let $origFileData := string(request:get-uploaded-file-data('file'))
			(: unzip erwartet base64Binary, die vom Upload direkt geliefert werden :)
			let $unpack := compression:unzip($origFileData, $filter, (), $entry-data, ())
			let $incoming := 
				<pack>{
					for $item in $unpack
						return $item}
				</pack>
			
			let $type := if (local-name($incoming/*[1]) = 'document-content')
				then 'OO'
				else 'MS'
				
			let $xslt := if ($type = 'OO')
				then doc('xslt/oo2tei.xsl')
				else doc('xslt/ms2tei.xsl')
				
			return transform:transform($incoming, $xslt, ())