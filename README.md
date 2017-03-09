# chefferie-de-projet
Some tools for project management / reporting / ...

## Usage
1. Create a slide template, containing all elements that you want to automatically update
2. Use the "Get-PowerpointContent" commandlet to analyze the slide and get the elements IDs
3. Modify the "Write-SlidesFromPowerpoint" commandlet to match your Excel file and the data to use
4. Run "Write-SlidesFromPowerpoint"

## Example
The following slide...
[slide template]: https://github.com/arnaudsoullie/chefferie-de-projet/sample_slide.jpg
gives the following result:
``powershell
PS C:\Users\Arnaud\Documents\chefferie-de-projet> Get-PowerpointContent -SlideTemplate C:\Users\Arnaud\Documents\example
.pptx
Slide title:  Title

### Slide content ###
Shape  2  -->  Title
Shape  3  -->  Subtitle
```
