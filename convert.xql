<html>
	<head>
		<title>Table Conversion</title>
	</head>
	<body>
		<div id="content">
			<form enctype="multipart/form-data" method="post" action="/apps/table2tei/convert2.xql">
				<fieldset>
					<legend>Upload Excel (.xlsx) or OO (.ods):</legend>
					<input type="file" name="file" style="width: 90%;"/>
					<input type="submit" value="Upload" style="float: right;"/>
					<br/>
					<label>1st line is heading (fixed rows in Excel are always considered headings): </label>
					<input type="checkbox" name="firstheading" />
					<br/>
					<label>select XSLT for further processing (if any): </label>
					<select name="xslt">
						<option selected="selected">none</option>&
					</select>
				</fieldset>
			</form>
		</div>
	</body>
</html>