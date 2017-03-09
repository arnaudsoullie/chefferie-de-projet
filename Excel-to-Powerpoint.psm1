<#
 .Synopsis
  Creates PowerPoint slides based on Excel data

 .Description
  Creates slides based on a PowerPoint template and data extracted from an Excel file

 .Parameter SlideTemplate
  The slide template to use.

 .Parameter ExcelSheet
  The Excel file to use as imput data.

 .Example
   # Analyzes a PowerPoint template file
   Analyze-PPT -SlideTemplate c:\User\file_path.pptx

 .Example
   # Creates a slide deck from an Excel sheet
   Create-Slides -SlideTemplate c:\User\file_path.pptx -ExcelSheet c:\User\file_path.xlsx

#>


function Get-PowerpointContent {
	param(
		[Parameter(mandatory=$true)] $SlideTemplate
		)
	# Use Office module
	Add-type -AssemblyName office

	# Open PowerPoint file
	$PowerPoint = New-Object -ComObject powerpoint.application
	$PowerPoint.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
	$Prez = $PowerPoint.Presentations.open($SlideTemplate)

	# Select slide (index starts at 1...)
	$slide2 = $Prez.Slides.item(1)
	Write-Host "Slide title: "$slide2.Shapes.title.TextFrame.TextRange.Text
	Write-Host ""

	# List slide content
	Write-Host "### Slide content ###"
	foreach ($shape in $slide2.Shapes) {
		If ($shape.HasTextFrame) {
			Write-Host "Shape "$shape.id" --> "$shape.TextFrame.TextRange.Text
		}
		Else {
			Write-Host "Shape "$shape.id" --> "$shape.name
		}
	}

	$Prez.Close()
}



Function Write-SlidesFromPowerpoint {
	param(
		[Parameter(mandatory=$true)] $SlideTemplate,
		[Parameter(mandatory=$true)] $ExcelSheet
		)
	# Use Office module
	Add-type -AssemblyName office

	# Read Excel file
	$Excel = New-Object -ComObject Excel.Application
	$workbook = $Excel.workbooks.open($ExcelSheet)
	$Excel.Visible= $false

	# Select a worksheet
	$worksheet = $workbook.worksheets.Item('projets')
	# Iterate on all cells from column A
	$range = $worksheet.UsedRange.rows.count
	$nb_projet = $range - 1
	$end_cell = "A"+$range
	Write-Host "Number of projects : "$nb_projet
	Write-Host ""
	$projets = $worksheet.Range("A2:$end_cell").Formula
	Write-Host "### Project list ###"


	# Get project's data from each tab
	Write-Host "Getting projects data ..."
	$projects_array = @()
	Foreach ($projet in $projets) {
		Write-host "   -- "$projet
		# Select the corresponding worksheet
		$worksheet = $workbook.worksheets.Item($projet)
		$cells = $worksheet.Cells

		# For each project, read some cell values
		# Of course, this is the part you need to adapt to your Excel file
		$project_hashtable = @{}
		$project_hashtable."name" = $cells.Item('1', 'B').text
		$project_hashtable."description"  = $cells.Item('2', 'B').text
		$project_hashtable."category" = $cells.Item('3', 'B').text
		$project_hashtable."people" = $cells.Item('4', 'B').text
		$project_hashtable."done" = $cells.Item('5', 'B').text
		$project_hashtable."todo" = $cells.Item('6', 'B').text
		$project_hashtable."comments" = $cells.Item('7', 'B').text
		$project_hashtable."workload_past" = $cells.Item('8', 'B').text
		$project_hashtable."workload_current" = $cells.Item('9', 'B').text
		$project_hashtable."weather" = $cells.Item('10', 'B').text
		#$project_hashtable
		$projects_array += $project_hashtable
	}

	$workbook.Close()

	# Add a slide for each project
	$PowerPoint = New-Object -ComObject powerpoint.application
	$PowerPoint.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
	$Prez = $PowerPoint.Presentations.open($SlideTemplate)
	$Slide = $Prez.Slides.item(1)


	Foreach ($project in $projects_array) {

		$newSlide  = $Slide.Duplicate()

		# Reuse here the shape ids identified using Get-PowerpointContent
		Foreach ($shape in $newSlide.shapes) {
			switch($shape.id) {
				4 {$shape.TextFrame.TextRange.Text = $project."description"}
				5 {$shape.TextFrame.TextRange.Text = $project."category"}
				6 {$shape.TextFrame.TextRange.Text = $project."name"}
				7 {$shape.TextFrame.TextRange.Text = $project."people"}
				8 {$shape.TextFrame.TextRange.Text = $project."done"}
				10 {$shape.TextFrame.TextRange.Text = $project."workload_current" + " (" + $project."workload_past" + ")"}
				12 {$shape.TextFrame.TextRange.Text = $project."todo"}
				13 {$shape.TextFrame.TextRange.Text = $project."comments"}
			}
		}
		$dirName = Get-Location
		# Add a picture based on the "weather" value
		$weather_picture = $project."weather" + ".png"
        $ImageFilename = Join-Path -Path $dirName -ChildPath $weather_picture
		Write-Host $ImageFilename
		$newSlide.Shapes.AddPicture($ImageFilename, $True, $True, 600, 20, 100, 100)
	}
}
